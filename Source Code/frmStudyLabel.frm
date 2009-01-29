VERSION 5.00
Begin VB.Form frmStudyLabel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Study Version"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   4860
      TabIndex        =   2
      Top             =   3180
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3660
      TabIndex        =   1
      Top             =   3180
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Caption         =   "New Study Version"
      Height          =   3075
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   5955
      Begin VB.TextBox txtStudyDescription 
         Height          =   1455
         Left            =   2220
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Label lblStudyVersion 
         Caption         =   "Version No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2220
         TabIndex        =   8
         Top             =   900
         Width           =   3195
      End
      Begin VB.Label lblStudyCode 
         Caption         =   "Study Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2220
         TabIndex        =   7
         Top             =   480
         Width           =   3195
      End
      Begin VB.Label Label8 
         Caption         =   "Study Change Description"
         Height          =   555
         Left            =   180
         TabIndex        =   6
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Study Version Number"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   900
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Study Code"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   480
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmStudyLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1998-2002. All Rights Reserved
'   File:       frmStudyLabel.frm
'   Author:     David Hook, 07/08/2002
'   Purpose:    Create a Study version for distribution to remote sites
'--------------------------------------------------------------------------------
'Revisions

'ZA 18/09/2002   Update List_sites.js/List_visits.js/List_forms.js/List_questions.js
'                if user creates a new site
' ic 25/11/2002  ic 25/11/2002 changed call arguements to create js list file functions
'                in cmdOK_click()
' DPH 28/01/2003 Generate HTML when creating a new version
' ic 13/02/2003 added mlClinicalTrialId arg to function calls
Option Explicit
Private mlClinicalTrialId As Long
Private msSelectedClinicalTrialName As String
Private mnClinicalTrialVersion As Integer
Private mnStudyVersioningNumber As Integer

'--------------------------------------------------------------------------------
Private Sub cmdCancel_Click()
'--------------------------------------------------------------------------------
' Unload the form
'--------------------------------------------------------------------------------

    Unload Me

End Sub

'--------------------------------------------------------------------------------
Private Sub cmdOK_Click()
'--------------------------------------------------------------------------------
' Create new Version by exporting study definition and copying to HTML folder
' revisions
' ic 25/11/2002 changed call arguements to create js list file functions
' DPH 28/01/2003 Generate HTML when creating a new version
' ic 13/02/2003 added mlClinicalTrialId arg to function calls
'--------------------------------------------------------------------------------

Dim oExchange As New clsExchange
Dim sCabFileName As String

On Error GoTo ErrHandler

    Screen.MousePointer = vbHourglass

    Set oExchange = New clsExchange

    '24/2/00, Note that ExportSDD now returns the name of the ceated Cab file (excluding path)
    sCabFileName = oExchange.ExportSDD(mlClinicalTrialId, msSelectedClinicalTrialName, _
                    mnClinicalTrialVersion, mnStudyVersioningNumber)

    ' DPH 17/10/2001 - Make sure cab file has been created
    If sCabFileName <> "" Then
        ' Do not write message to message table unless successfully copied

        'Mo Morris 24/2/00
        Do Until FileExists(gsOUT_FOLDER_LOCATION & sCabFileName)
            DoEvents
        Loop

        On Error GoTo CopyFileErr

        'Mo Morris 24/2/00
        Call FileCopy(gsOUT_FOLDER_LOCATION & sCabFileName, _
                        goUser.Database.HTMLLocation & sCabFileName)

        On Error GoTo ErrHandler

        ' DPH 07/08/2002 - No longer distributes automatically
        ' update StudyVersion table
        Call UpdateStudyVersionNumber
        
    Else
        Call DialogError("Create new version aborted - unable to create export file", "Study Version Creation")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    ' DPH 28/01/2003 Generate HTML when creating a new version (also user list)
    Call PublishStudy(mlClinicalTrialId, gnCurrentVersionId(mlClinicalTrialId))
    Call CreateUsersList(goUser)
    
    'ic 13/02/2003 added mlClinicalTrialId arg to function calls
    'ZA 18/09/2002 - create List_questions.js file with question details
    Call CreateQuestionsList(goUser, mlClinicalTrialId)
    'Creates List_Studies.js file with list of trials
    'CreateTrialList
    'creates List_visits.js file with the the list of visits
    Call CreateVisitList(goUser, mlClinicalTrialId)
    'creates List_forms.js file with the list of all eForms
    Call CreateEFormsList(goUser, mlClinicalTrialId)
    
    Screen.MousePointer = vbDefault
    
    Call DialogInformation("New version of " & msSelectedClinicalTrialName & " has successfully been created", "Study Version Creation")
    
    cmdOK.Enabled = False
    cmdCancel.Caption = "&Close"

Exit Sub
CopyFileErr:
    ' If an error copying file
    Screen.MousePointer = vbDefault

    Call DialogError("Distribute new version aborted - Error copying file to published HTML folder", "Study Definition Distribution")

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "cmdOK_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

Private Sub Form_Load()
'--------------------------------------------------------------------------------
' Load form setting up all labels
'--------------------------------------------------------------------------------

    Me.Icon = frmMenu.Icon

    ' Set labels + initialise description textbox
    lblStudyCode = msSelectedClinicalTrialName
    mnStudyVersioningNumber = NextStudyVersion
    lblStudyVersion = mnStudyVersioningNumber
    txtStudyDescription = ""

End Sub

Public Sub InitialiseMe(lClinicalTrialId As Long, sSelectedClinicalTrialName As String, nClinicalTrialVersion As Integer)
'--------------------------------------------------------------------------------
' Get required information & store in form
'--------------------------------------------------------------------------------

    mlClinicalTrialId = lClinicalTrialId
    msSelectedClinicalTrialName = sSelectedClinicalTrialName
    mnClinicalTrialVersion = nClinicalTrialVersion

End Sub

'---------------------------------------------------------------------
Private Function NextStudyVersion() As Integer
'--------------------------------------------------------------------------------
' Get the next study version number to distribute to remote sites
' TODO : Need to make this multiuser safe
'--------------------------------------------------------------------------------

Dim rsNextStudyVersion As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler

    sSQL = "SELECT max(StudyVersion) as CurrentVersionId FROM StudyVersion " _
        & " WHERE ClinicalTrialId = " & mlClinicalTrialId
    Set rsNextStudyVersion = New ADODB.Recordset
    rsNextStudyVersion.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

    If IsNull(rsNextStudyVersion!CurrentVersionId) Then
        NextStudyVersion = 1
    Else
        NextStudyVersion = rsNextStudyVersion!CurrentVersionId + 1
    End If

    rsNextStudyVersion.Close
    Set rsNextStudyVersion = Nothing

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                        "NextStudyVersion", "frmStudyLabel")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function

'--------------------------------------------------------------------------------
Private Function UpdateStudyVersionNumber()
'--------------------------------------------------------------------------------
' Update the Study Version table with labelling details
'--------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    Dim sDescription As String
    Dim sSQL As String

    ' Setup description
    sDescription = txtStudyDescription.Text
    If sDescription = "" Then
        sDescription = "NULL"
    Else
        sDescription = "'" & Replace(sDescription, "'", "") & "'"
    End If

    'TA 08/04/2003: use sqlstandardnow
    sSQL = "INSERT INTO StudyVersion (ClinicalTrialId, StudyVersion, VersionTimeStamp, VersionDescription) " _
        & " VALUES ( " & mlClinicalTrialId & "," & mnStudyVersioningNumber & "," _
        & SQLStandardNow & "," & sDescription & ")"

    ' Execute SQL
    MacroADODBConnection.Execute sSQL

Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                        "UpdateStudyVersionNumber", "frmStudyLabel")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function
