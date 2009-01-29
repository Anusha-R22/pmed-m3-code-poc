VERSION 5.00
Begin VB.Form frmExportPatientData 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export Subjects"
   ClientHeight    =   3450
   ClientLeft      =   555
   ClientTop       =   4680
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3450
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExportSubject 
      Caption         =   "&Export"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   "Subject"
      Height          =   675
      Left            =   1440
      TabIndex        =   10
      Top             =   1320
      Width           =   1395
      Begin VB.ComboBox cboSubject 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Study"
      Height          =   675
      Left            =   60
      TabIndex        =   9
      Top             =   540
      Width           =   2775
      Begin VB.ComboBox cboStudy 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2565
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Site"
      Height          =   675
      Left            =   60
      TabIndex        =   8
      Top             =   1320
      Width           =   1335
      Begin VB.ComboBox cboSite 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Export Type"
      Height          =   795
      Left            =   60
      TabIndex        =   7
      Top             =   2100
      Width           =   2775
      Begin VB.OptionButton optAllData 
         Caption         =   "All data"
         Height          =   225
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1020
      End
      Begin VB.OptionButton optChangedData 
         Caption         =   "Data marked as 'Changed'"
         Height          =   225
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2205
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Select a study, site (optional), and subject (optional)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   2775
   End
End
Attribute VB_Name = "frmExportPatientData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       frmExportPatientData.frm
'   Author:     Andrew Newbigging, June 1997
'   Purpose:    Allows selection of patient data to be exported, either all data
'   or data changed since last export
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
'   Revisions:
'   1    Mo Morris   26/01/98
'   2    Mo Morris   24/02/98
'   3    Andrewn     2/04/98
'   4    PN          10/09/99 Upgrade from DAO to ADO and updated code to conform
'                             to VB standards doc version 1.0
'   5    PN          15/09/99 Changed call to ADODBConnection() to MacroADODBConnection()
'   PN  21/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   WillC 10/11/99  Added error handlers
'   TA 08/05/2000   subclassing removed
'Mo Morris  2/2/01
'   Changes made so that the export can filter on Study, Site and PersonId
'   Unused controls ProgressBar1,txtCurrentFile,Label2 and Label3 removed
' NCJ/TA 7 Feb 01 - Minor wording changes on labels
'---------------------------------------------------------------------
Option Explicit
Option Base 0
Option Compare Binary
'---------------------------------------------------------------------

Private glSelectedTrialId As Long
Private gsSelectedSite As String
Private gsSelectedPersonId As String
'ASH 11/12/2002
Private msDatabase As String
Private oDatabase As MACROUserBS30.Database
Private bLoad As Boolean
Private sConnectionString As String
Private sMessage As String
Private mconMACRO As ADODB.Connection


'---------------------------------------------------------------------
Private Function ExportPRD(b As Boolean) As String
'---------------------------------------------------------------------
' this function will run the export
'---------------------------------------------------------------------
Dim oExchange As clsExchange

    On Error GoTo ErrHandler

    Screen.MousePointer = vbHourglass
    Set oExchange = New clsExchange
    ExportPRD = oExchange.ExportPRD(b, glSelectedTrialId, gsSelectedSite, gsSelectedPersonId)
    Set oExchange = Nothing
    Screen.MousePointer = vbDefault

Exit Function
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "ExportPRD")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select

End Function

'---------------------------------------------------------------------
Private Sub cboSite_Click()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    gsSelectedSite = cboSite.Text
    
    cboSubject.Enabled = True
    LoadPersonCombo

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cboSite_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub cboStudy_Click()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    glSelectedTrialId = cboStudy.ItemData(cboStudy.ListIndex)
    
    cmdExportSubject.Enabled = True
    
    'enable cmdStartAll or cmdStartChanged
    If optChangedData.Value = True Then
   '     cmdStartChanged.Enabled = True
    Else
   '     cmdStartAll.Enabled = True
    End If
    
    cboSite.Enabled = True
    LoadSiteCombo

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cboStudy_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub cboSubject_Click()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    gsSelectedPersonId = cboSubject.Text

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cboSubject_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub cmdExportSubject_Click()
'---------------------------------------------------------------------
Dim sError As String

    On Error GoTo ErrHandler

     sError = ExportPRD(optChangedData.Value)

    If sError = "" Then
        DialogInformation "Export successful", "Subject Export"
    Else
        DialogError sError, "Subject Export"
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdExportSubject_Click")
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

    On Error GoTo ErrHandler

    Me.Icon = frmMenu.Icon

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "Form_Load")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub LoadStudyCombo()
'---------------------------------------------------------------------
Dim rsTrial As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler

    cboStudy.Clear

    sSQL = "SELECT ClinicalTrialId, ClinicalTrialName " _
        & " FROM ClinicalTrial " _
        & " WHERE ClinicalTrialId  > 0 " _
        & " ORDER BY ClinicalTrialName"
    
    Set rsTrial = New ADODB.Recordset
    rsTrial.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    While Not rsTrial.EOF
        cboStudy.AddItem rsTrial![ClinicalTrialName]
        cboStudy.ItemData(cboStudy.NewIndex) = rsTrial![ClinicalTrialId]
        rsTrial.MoveNext
    Wend
    
    rsTrial.Close
    Set rsTrial = Nothing

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "LoadStudyCombo")
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
Private Sub LoadSiteCombo()
'---------------------------------------------------------------------
Dim rsSite As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler

    cboSite.Clear
    
    cboSite.AddItem "All Sites", 0
    
    sSQL = "SELECT DISTINCT TrialSite " _
        & " FROM TrialSubject " _
        & " WHERE ClinicalTrialId = " & glSelectedTrialId _
        & " ORDER BY TrialSite"
        
    Set rsSite = New ADODB.Recordset
    rsSite.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    While Not rsSite.EOF
        cboSite.AddItem rsSite![TrialSite]
        rsSite.MoveNext
    Wend

    rsSite.Close
    Set rsSite = Nothing
    
    cboSite.ListIndex = 0
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "LoadSiteCombo")
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
Private Sub LoadPersonCombo()
'---------------------------------------------------------------------
Dim rsPerson As ADODB.Recordset
Dim sSQL As String
Dim i As Integer

    On Error GoTo ErrHandler

    cboSubject.Clear

    cboSubject.AddItem "All Subjects", 0
    
    sSQL = "SELECT max(PersonID) as MaxPersonId FROM TrialSubject" _
        & " WHERE ClinicalTrialId = " & glSelectedTrialId _
        & " AND TrialSite = '" & gsSelectedSite & "'"
        
    Set rsPerson = New ADODB.Recordset
    rsPerson.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If Not IsNull(rsPerson!MaxPersonId) Then
        For i = 1 To rsPerson!MaxPersonId
            cboSubject.AddItem i
            cboSubject.ItemData(cboSubject.NewIndex) = i
        Next
    End If
    rsPerson.Close
    Set rsPerson = Nothing
    cboSubject.ListIndex = 0

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "LoadPersonCombo")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'---------------------------------------------------------------------------
Public Sub Display(Optional ByVal sDatabase As String)
'---------------------------------------------------------------------------
'
'---------------------------------------------------------------------------
'Dim msDatabase As String
'
'    msDatabase = sDatabase
'    Set oDatabase = New MACROUserBS30.Database
'    bLoad = oDatabase.Load(SecurityADODBConnection, goUser.UserName, msDatabase, "", False, sMessage)
'    sConnectionString = oDatabase.ConnectionString
'    Set mconMACRO = New ADODB.Connection
'    mconMACRO.Open sConnectionString
'    mconMACRO.CursorLocation = adUseClient
    
    cboSite.Enabled = False
    cboSubject.Enabled = False
    
    LoadStudyCombo
    
    'TA 02/04/2001: if there is at least one study then select the first one
    If cboStudy.ListCount > 0 Then
        cboStudy.ListIndex = 0
        cmdExportSubject.Enabled = True
    Else
        cmdExportSubject.Enabled = False
    End If
    
    optChangedData.Value = False
    optAllData.Value = True
    
    Me.Caption = "Export Subjects " & "[" & goUser.DatabaseCode & "]"
    Me.Show vbModal

End Sub
