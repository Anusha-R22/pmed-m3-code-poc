VERSION 5.00
Begin VB.Form frmExportStudyDefinition 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export Study Definition"
   ClientHeight    =   1455
   ClientLeft      =   555
   ClientTop       =   4680
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1455
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboStudies 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   4575
   End
   Begin VB.CommandButton cmdStartExport 
      Caption         =   "Start Export"
      Height          =   372
      Left            =   3240
      TabIndex        =   1
      Top             =   960
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "Start by selecting Study Definition to be Exported"
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4572
   End
End
Attribute VB_Name = "frmExportStudyDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       frmExportStudyDefinition.frm
'   Author:         Mo Morris July 1997
'   Purpose:    Allows selection of study to be exported.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
 '  1   Mo Morris           4/07/97
'   2   Mo Morris           4/07/97
'   3   Andrew Newbigging   11/07/97
'   4   Mo Morris           14/08/97
'   5   Andrew Newbigging   10/09/97
'   6   Andrew Newbigging   18/09/97
'   7   Mo Morris           18/09/97
'   8   Mo Morris           26/09/97
'   9   Mo Morris           24/10/97
'   10  Andrew Newbigging   2/12/97
'   11  Mo Morris           24/02/98
'   12  Andrew Newbigging   2/04/98
'   13  Andrew Newbigging   30/04/98
'   14  Joanne Lau          8/05/98
'       Mo Morris           13/10/98    SPR 544
'                                       cmdSelectExportFile_Click changed so that it correctly
'                                       detects a Cancel from the CommonDialog.ShowOpen
'   15  PN  14/09/99    Updated code to conform to VB standards doc version 1.0
'                       Upgrade database access code from DAO to ADO
'   16  PN  15/09/99    Changed call to ADODBConnection() to MacroADODBConnection()
'   PN  21/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   Willc 10/11/99  Added the error handlers
'   Mo Morris 6/12/99   Data Control and Date Grid changed to ADO enabled controls
'   Mo Morris   16/2/00 Data Control and Data Grid replaced by Combo
'   Mo Morris   23/2/00 Major re-write. Export file name selection removed.
'                       Export file name is no longer passed to ExportSDD.
'                       cmdSelectExportFile, ProgressBar1, CommonDialog1 and  Label2 removed
'   TA 08/05/2000   subclassing removed
' MLM 23/04/07: Bug 2762: cboStudies.Sorted = True
'----------------------------------------------------------------------------------------'
Option Explicit
Option Compare Binary
Option Base 0

Private msSelTrialId As String
Private msSelTrialName As String
Private msSelVersionId As String
'ASH 11/12/2002
Private oDatabase As MACROUserBS30.Database
Private bLoad As Boolean
Private sConnectionString As String
Private sMessage As String
Private mconMACRO As ADODB.Connection
Private msDatabase As String


'---------------------------------------------------------------------
Private Sub cboStudies_Click()
'---------------------------------------------------------------------
'Changed Mo Morris 7/4/00 SR 3315, exit if Listindex=-1, error trapping added
'---------------------------------------------------------------------
On Error GoTo ErrHandler

    If cboStudies.ListIndex = -1 Then
        cmdStartExport.Enabled = False
        Exit Sub
    Else
        cmdStartExport.Enabled = True
    End If
    
    msSelTrialId = cboStudies.ItemData(cboStudies.ListIndex)
    msSelTrialName = Mid(cboStudies.Text, 1, InStr(cboStudies.Text, " - ") - 1)
    'changed Mo Morris 28/4/00, SR3319, instead of hardcoding version to 1 call gnCurrentVersionId
    'msSelVersionId = "1"
    msSelVersionId = gnCurrentVersionId(msSelTrialId)

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "cboStudies_Click")
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

Dim oExchange As clsExchange
Dim sFile As String
    On Error GoTo ErrHandler
    HourglassOn
    
    Set oExchange = New clsExchange
    
    ' DPH 17/10/2001 - Check for file creation 22/10/2001 - Removed <> "" Check
    sFile = oExchange.ExportSDD(msSelTrialId, msSelTrialName, msSelVersionId) '<> ""
    
'    oExchange.ExportStudyCAB msSelTrialId, msSelVersionId
'    oExchange.ExportSDD msSelTrialId, msSelTrialName, msSelVersionId, msExportFile
'    oExchange.ExportStudyCAB msSelTrialId, msSelVersionId, msExportFile
    
    cmdStartExport.Enabled = False
    
    HourglassOff
    
    'Mo Morris 12/12/00 / DPH 17/10/2001
    If sFile <> "" Then
        Call DialogInformation(msSelTrialName & " has successfully been exported", "Study Definition Export")
    Else
        Call DialogError("Export failed as could not create file", "Study Definition Export")
    End If

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
'Mo Morris 16/2/00 code to load grdStudyDefinitions replaced by code to load cboStudies
'---------------------------------------------------------------------
    
    Me.Icon = frmMenu.Icon
    
End Sub

'--------------------------------------------------------------------
Public Sub Display(Optional ByVal sDatabase As String)
'--------------------------------------------------------------------
'
'--------------------------------------------------------------------
Dim sSQL As String
Dim rsStudies As ADODB.Recordset
Dim sConToUse As ADODB.Connection

    On Error GoTo ErrHandler
    msDatabase = sDatabase
    cmdStartExport.Enabled = False
    If msDatabase <> "" Then
        Set oDatabase = New MACROUserBS30.Database
        bLoad = oDatabase.Load(SecurityADODBConnection, goUser.UserName, msDatabase, "", False, sMessage)
        sConnectionString = oDatabase.ConnectionString
        Set mconMACRO = New ADODB.Connection
        mconMACRO.Open sConnectionString
        mconMACRO.CursorLocation = adUseClient
        Set sConToUse = mconMACRO
    Else
        Set sConToUse = MacroADODBConnection
    End If

    
    sSQL = "SELECT ClinicalTrial.ClinicalTrialId, ClinicalTrial.ClinicalTrialName,"
    sSQL = sSQL & " ClinicalTrial.ClinicalTrialDescription, StudyDefinition.VersionId "
    sSQL = sSQL & " FROM ClinicalTrial, StudyDefinition "
    sSQL = sSQL & " WHERE ClinicalTrial.ClinicalTrialId = StudyDefinition.ClinicalTrialId"
    Set rsStudies = New ADODB.Recordset
    rsStudies.Open sSQL, sConToUse, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    cboStudies.Clear
    Do Until rsStudies.EOF
        cboStudies.AddItem rsStudies!ClinicalTrialName & " - " & rsStudies!ClinicalTrialDescription
        cboStudies.ItemData(cboStudies.NewIndex) = rsStudies!ClinicalTrialId
        rsStudies.MoveNext
    Loop
    
    Me.Caption = "Export Study Definition " & "[" & goUser.DatabaseCode & "]"
    Me.Show vbModal
    
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
