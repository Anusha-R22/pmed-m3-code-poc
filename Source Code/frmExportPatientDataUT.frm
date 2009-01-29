VERSION 5.00
Begin VB.Form frmExportPatientData 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export Subjects"
   ClientHeight    =   3480
   ClientLeft      =   555
   ClientTop       =   4680
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3480
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   3000
      Width           =   1155
   End
   Begin VB.CommandButton cmdExportSubject 
      Caption         =   "&Export"
      Height          =   375
      Left            =   60
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   "Subject"
      Height          =   675
      Left            =   1710
      TabIndex        =   10
      Top             =   1320
      Width           =   1600
      Begin VB.ComboBox cboSubject 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Study"
      Height          =   675
      Left            =   60
      TabIndex        =   9
      Top             =   540
      Width           =   3255
      Begin VB.ComboBox cboStudy 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2985
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Site"
      Height          =   675
      Left            =   60
      TabIndex        =   8
      Top             =   1320
      Width           =   1600
      Begin VB.ComboBox cboSite 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Patient Data Export Type"
      Height          =   795
      Left            =   60
      TabIndex        =   7
      Top             =   2100
      Width           =   3255
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
      Width           =   3255
   End
End
Attribute VB_Name = "frmExportPatientData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
' File:         frmExportPatientDataUT.frm
' Copyright:    InferMed Ltd. 2003. All Rights Reserved
' Author:       Richard Meinesz September 2003
' Purpose:      Allows selection of patient data to be exported
'               This is a version of frmExportPatientData.frm adapted for MACRO Utilities
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
'   Revisions:
'   REM 12 Sept 03 - Created from copy of frmExportPatientData.frm
'   NCJ 29 Oct 03 - Changed file header and comments
'---------------------------------------------------------------------
Option Explicit
Option Base 0
Option Compare Binary
'---------------------------------------------------------------------

Private glSelectedTrialId As Long
Private gsSelectedTrialName As String
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
Private Function ExportPRD(bChangedData As Boolean) As String
'---------------------------------------------------------------------
' this function will run the export
'---------------------------------------------------------------------
Dim oExchange As clsExchange
Dim sZipFileName As String

    On Error GoTo ErrHandler
    
    'Export patient data
    Screen.MousePointer = vbHourglass
    Set oExchange = New clsExchange
    ExportPRD = oExchange.ExportPRD(bChangedData, glSelectedTrialId, gsSelectedSite, gsSelectedPersonId, sZipFileName)
    Set oExchange = Nothing
    
    If ExportPRD = "" Then
        'Export MIMessage, LFMessage data
        ExportPRD = ExportMIMessagesLFMessages(gsSelectedTrialName, gsSelectedSite, gsSelectedPersonId, sZipFileName)
    End If
    Screen.MousePointer = vbDefault

Exit Function
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "ExportPRD")
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
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub cboStudy_Click()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    glSelectedTrialId = cboStudy.ItemData(cboStudy.ListIndex)
    gsSelectedTrialName = cboStudy.Text
    
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
            Call ExitMACRO
            Call MACROEnd
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
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub cmdClose_Click()
'---------------------------------------------------------------------
    
    Unload Me

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
            Call ExitMACRO
            Call MACROEnd
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
            Call ExitMACRO
            Call MACROEnd
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
        cboStudy.AddItem rsTrial![ClinicaltrialName]
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
Public Sub Display()
'---------------------------------------------------------------------------
'
'---------------------------------------------------------------------------

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
