VERSION 5.00
Begin VB.Form frmGenerateHTML 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generate HTML Forms"
   ClientHeight    =   1290
   ClientLeft      =   13260
   ClientTop       =   10050
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Height          =   375
      Left            =   2820
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select study definition"
      Height          =   675
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3975
      Begin VB.ComboBox cboTrialList 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   3675
      End
   End
End
Attribute VB_Name = "frmGenerateHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1998-2004. All Rights Reserved
'   File:       frmGenerateHTML.frm
'   Author:     Andrew Newbigging, Febuary 1998
'   Purpose:    Allows selection of study to generate HTML versions of CRF pages.
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
'   Revisions:
'   1    Andrew Newbigging   24/02/98
'   2    Andrew Newbigging   24/02/98
'   3    Joanne Lau          5/03/98
'   4    Joanne Lau         13/03/98
'   PN  21/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   WillC 10/11/99  Added the error handlers
'   Mo 13/12/99     Id's from integer to Long
'   WillC 2/3/2000  Changed from grid to combobox SR3054.
'   TA 08/05/2000   subclassing removed
'   TA 19/10/2000: Use new hourglass functions to ensure hourglass displayed
'   JL 8/3/01   : Commented out CreateNetscapefiles (in cmdGenerate) as we have stopped
'                 supporting netscape for now.
'   NCJ 15 Jul 04 - Bug 2342 - Exclude Library from list of studies
'--------------------------------------------------------------------------------

Option Explicit
Option Base 0
Option Compare Binary

'MLM 20/06/05: These are no longer appropriate with multi-study generation
'Private mlSelTrialId As Long
'Private msSelTrialName As String
'Private msSelVersionId As String

'---------------------------------------------------------------------
Private Sub cboTrialList_Click()
'---------------------------------------------------------------------
' enable the generate button if something is chosen
'---------------------------------------------------------------------
On Error GoTo ErrHandler
    
    If cboTrialList.ListIndex = -1 Then
        cmdGenerate.Enabled = False
        'changed Mo Morris 7/4/00 SR 3315, exit sub added
        Exit Sub
    Else
        cmdGenerate.Enabled = True
    End If
    
    'mlSelTrialId = cboTrialList.ItemData(cboTrialList.ListIndex)
    'msSelTrialName = Trim(cboTrialList.Text)
    'changed Mo Morris 28/4/00, SR3319, instead of hardcoding version to 1 call gnCurrentVersionId
    'msSelVersionId = "1"
    'msSelVersionId = gnCurrentVersionId(mlSelTrialId)

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
Private Sub cmdClose_Click()
'---------------------------------------------------------------------
'Close form
'---------------------------------------------------------------------

    Unload Me
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdGenerate_Click()
'---------------------------------------------------------------------
' Generate the HTML
' MLM 20/06/05: bug 2528: Added the "all studies" options
'---------------------------------------------------------------------
Dim lCount As Long

    On Error GoTo ErrHandler
    If cboTrialList.ItemData(cboTrialList.ListIndex) = 0 Then
        'all studies
        HourglassOn
        For lCount = 1 To cboTrialList.ListCount - 1
            GenerateHtml cboTrialList.ItemData(lCount)
        Next lCount
        HourglassOff
        DialogInformation "All studies generated successfully."
    Else
        'only the selected study
        HourglassOn
        GenerateHtml cboTrialList.ItemData(cboTrialList.ListIndex)
        HourglassOff
        DialogInformation "Study " & cboTrialList.Text & " generated successfully."
    End If
    
    Exit Sub
ErrHandler:
      Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                    "cmdGenerate_Click")
            Case OnErrorAction.Ignore
                Resume Next
            Case OnErrorAction.Retry
                Resume
            Case OnErrorAction.QuitMACRO
                Unload frmMenu
       End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub GenerateHtml(lStudyId As Long)
'---------------------------------------------------------------------
'MLM 20/06/05: bug 2528: created: generate the HTML for the specified study.
'---------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    'ic 02/04/2002 function modHTML/CreateHTMLFiles() renamed PublishEformASPFiles()
    'ic 29/10/2002 changed to create static html - eforms now created on-the-fly
    'PublishEformASPFiles mlSelTrialId, gnCurrentVersionId(mlSelTrialId)
    Call PublishStudy(lStudyId, gnCurrentVersionId(lStudyId))
    'CreateNetscapefiles mlSelTrialId, gnCurrentVersionId(mlSelTrialId)
    'ic 11/02/2003, added trialid to calls
    Call CreateVisitList(goUser, lStudyId)
    Call CreateEFormsList(goUser, lStudyId)
    Call CreateQuestionsList(goUser, lStudyId)
    Call CreateUsersList(goUser)
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmGenerateHTML.GenerateHtml(" & lStudyId & ")"
End Sub

'--------------------------------------------------------------------------------
Private Sub Form_Load()
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    Me.Icon = frmMenu.Icon

    FormCentre Me

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

'---------------------------------------------------------------------------
Public Sub Display()
'---------------------------------------------------------------------------
'Loads combo
' NCJ 15 Jul 04 - Bug 2342 - Exclude Library from list
'---------------------------------------------------------------------------
Dim sSQL As String
Dim rsTrials As ADODB.Recordset
    
    cmdGenerate.Enabled = False

    Set rsTrials = New ADODB.Recordset

    ' NCJ 15 Jul 04 - Added WHERE ClinicalTrialId > 0
    sSQL = "SELECT * FROM ClinicalTrial WHERE ClinicalTrialId > 0"
    
    rsTrials.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    cboTrialList.Clear
    
    ' MLM 20/06/05: bug 2528: offer an "all studies" option:
    If rsTrials.RecordCount > 1 Then
        cboTrialList.AddItem "All Studies"
    End If
    
    Do Until rsTrials.EOF
        cboTrialList.AddItem rsTrials!ClinicalTrialName '& " - " & rsTrials!ClinicalTrialDescription
        cboTrialList.ItemData(cboTrialList.NewIndex) = rsTrials!ClinicalTrialId
        rsTrials.MoveNext
    Loop
    
    rsTrials.Close
    Set rsTrials = Nothing
    
    Me.Caption = "Generate HTML Forms " & "[" & goUser.DatabaseCode & "]"
    Me.Show vbModal

End Sub
