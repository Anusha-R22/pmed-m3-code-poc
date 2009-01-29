VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImportStudyDefinition 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Study Definition"
   ClientHeight    =   975
   ClientLeft      =   1245
   ClientTop       =   2280
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   975
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStartImport 
      Caption         =   "Start Import"
      Height          =   372
      Left            =   3240
      TabIndex        =   2
      Top             =   480
      Width           =   1452
   End
   Begin VB.CommandButton cmdSelectImportFile 
      Caption         =   "Select Name/Location of Import File"
      Height          =   372
      Left            =   105
      TabIndex        =   1
      Top             =   480
      Width           =   3012
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   372
      Left            =   6120
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   2532
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5400
      Top             =   0
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Import Progress:"
      Height          =   252
      Left            =   4800
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "Start by selecting file to be Imported"
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4572
   End
End
Attribute VB_Name = "frmImportStudyDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       frmImportStudyDefinition.frm
'   Author:         Andrew Newbigging June 1997
'   Purpose:    Allows selection of file containing study definition to be imported.
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
'   9   Mo Morris           3/10/97
'   10  Mo Morris           24/10/97
'   11  Andrew Newbigging   2/12/97
'   12  Mo Morris           24/02/98
'   13  Andrew Newbigging   2/04/98
'   14  Mo Morris           13/05/98
'   15  Joanne Lau          15/05/98
'   16  Joanne Lau          15/05/98
'   17  Joanne Lau          19/05/98
'   18  Joanne Lau          15/06/98
'   19  Joanne Lau          15/06/98
'   20  Joanne Lau          18/06/98
'   21  Andrew Newbigging   11/11/98
'       Modified to do COMPACT import as well as standard MACRO import.
'       ImportType property is used to indicate which.
'   22  PN  10/09/99    Upgrade from DAO to ADO and updated code to conform
'                       to VB standards doc version 1.0
'   23  PN  15/09/99    Changed call to ADODBConnection() to MacroADODBConnection()
'   PN  21/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   Mo Morris   23/2/00 Changes made cmdStartImport_Click
'   TA 08/05/2000   subclassing removed
'   DPH 18/04/2002 - Include ZIP files
'------------------------------------------------------------------------------------'
'---------------------------------------------------------------------
Option Explicit
Option Compare Binary
Option Base 0
'---------------------------------------------------------------------

Private msImportFile As String
Private mnSDDImportType As Integer
'   ATN 11/11/98
'   New variable for name of trial imported from COMPACT
Private msClinicalTrialName As String
'ASH 11/12/2002
Private oDatabase As MACROUserBS30.Database
Private bLoad As Boolean
Private sConnectionString As String
Private sMessage As String
Private mconMACRO As ADODB.Connection
Private msDatabase As String

'---------------------------------------------------------------------
Public Property Get ImportType() As Integer
'---------------------------------------------------------------------

    ImportType = mnSDDImportType

End Property

'---------------------------------------------------------------------
Public Property Let ImportType(ByVal vNewValue As Integer)
'---------------------------------------------------------------------
'changed Mo Morris 20/3/00, SR 3262, diolog filter changed from '*.1;*.cab' to '*.cab'
' DPH 18/04/2002 - Include ZIP files
'---------------------------------------------------------------------
  
    mnSDDImportType = vNewValue
    
    If mnSDDImportType = SDDImportType.MACRO Then
        With CommonDialog1
            .DialogTitle = "Import File Selection"
            .InitDir = gsIN_FOLDER_LOCATION
            .DefaultExt = "cab"
            .Filter = "Study Definitions (*.cab;*.zip)|*.cab;*.zip"
        End With
    End If

End Property

'---------------------------------------------------------------------
Private Sub cmdSelectImportFile_Click()
'---------------------------------------------------------------------
' DPH 18/04/2002 - Include ZIP files
'---------------------------------------------------------------------
    
    cmdStartImport.Enabled = False
    CommonDialog1.CancelError = True
    
    On Error GoTo ErrInShowOpen
    CommonDialog1.ShowOpen
    msImportFile = CommonDialog1.FileName
    
    'MRCCOMPACT code taken out by Mo Morris 9/9/99
    
    'Mo 24/1/00, file extention check added
    'Mo 23/2/00 changed from .sdd to .cab
    If LCase(Mid(msImportFile, Len(msImportFile) - 3, 4)) <> ".cab" And _
        LCase(Mid(msImportFile, Len(msImportFile) - 3, 4)) <> ".zip" Then
        MsgBox ("Study Definition files must have an 'cab' or 'zip' extension")
        Exit Sub
    End If
    
    cmdStartImport.Enabled = True
    Exit Sub
    
ErrInShowOpen:
    
    
    If Err.Number <> 32755 Then
        MsgBox ("Unknown error during opening and checking of import file." _
            & Chr(13) & "Error code " & Err.Number & " - " & Err.Description _
            & Chr(13) & "Import Aborted.")
    End If
    
    cmdStartImport.Enabled = False

End Sub

'---------------------------------------------------------------------
Private Sub cmdStartImport_Click()
'---------------------------------------------------------------------
'Re-written by Mo Morris 23/2/00
'Study definition import is now basedd on a cab file (not an sdd file)
'call to ImportStudyDefinitionCAB added
' NCJ 21 Aug 06 - ImportSDD can now return "Study Locked" error (Bug 2642)
'---------------------------------------------------------------------
Dim oExchange As clsExchange
Dim sImportFile As String
Dim sDlgTitle As String

    On Error GoTo ErrHandler
    
    HourglassOn
    
    Set oExchange = New clsExchange
    
    'Unpack the CAB file into an SDD file, and individual study documents and
    'graphic files and place them in directory AppPath/CabExtract.
    Call oExchange.ImportStudyDefinitionCAB(msImportFile)
    
    'get the name of the SDD file from the CABExtract folder
    sImportFile = Dir(gsCAB_EXTRACT_LOCATION & "*.sdd")
    
    sDlgTitle = "Import Study Definition"
    'Start the import
    Select Case oExchange.ImportSDD(gsCAB_EXTRACT_LOCATION & sImportFile)
    Case ExchangeError.Success
        Call MsgBox("Import successfully completed.", , sDlgTitle)
    Case ExchangeError.EmptyFile
        Call MsgBox(msImportFile & " is empty." + vbNewLine + "Import aborted.", , sDlgTitle)
    Case ExchangeError.Invalid
        Call MsgBox(msImportFile & " is not a valid study definition." + vbNewLine + "Import aborted.", , sDlgTitle)
    Case ExchangeError.UserAborted
        Call MsgBox("Import aborted due to user intervention.", , sDlgTitle)
    ' DPH 17/10/2001 - New Error Case
    Case ExchangeError.DirectoryNotFound
        Call MsgBox("Import aborted due to absent file/folder.", , sDlgTitle)
    'Mo 21/11/2005 COD0210
    Case ExchangeError.MoreRecentVersion
        Call MsgBox("Import of a study from a more recent version of MACRO disallowed.", , sDlgTitle)
    ' NCJ 21 Aug 06 - Check if study locked
    Case ExchangeError.TrialLocked
        Call DialogInformation("Cannot import because the study is in use.", sDlgTitle)
    Case Else
        Call MsgBox("Unexpected error. Import aborted.", , sDlgTitle)
    End Select
    
    Set oExchange = Nothing
    HourglassOff
    cmdStartImport.Enabled = False
    
    'NOTE THAT THE CAB FILE IS NOT DELETED

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "cmdStartImport_Click")
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
'---------------------------------------------------------------------

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

'--------------------------------------------------------------
Public Sub Display(Optional ByVal sDatabase As String = "")
'--------------------------------------------------------------
'
'--------------------------------------------------------------
Dim sSQL As String
Dim rsExportImport As ADODB.Recordset
Dim sConToUse As ADODB.Connection
  
    On Error GoTo ErrHandler
    msDatabase = sDatabase
    cmdStartImport.Enabled = False
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

    
    'Initialise the progress bar's min and max settings based on
    'the number of files involved in the Export Parameters file
    ProgressBar1.Min = 0
    'Changed Mo Morris 13/10/00, change from table SDDExportImport to MACROTable
    sSQL = "Select * From MACROTable WHERE STYDEF = 1"
    
    ' PN 10/09/99  - upgrade to ado
    Set rsExportImport = New ADODB.Recordset
    rsExportImport.Open sSQL, sConToUse, adOpenKeyset, adLockReadOnly, adCmdText
    
    With rsExportImport
        .MoveLast
        ProgressBar1.Max = .RecordCount
        .Close
    End With
    
    Set rsExportImport = Nothing
    
    Me.Caption = "Import Study Definition " & "[" & goUser.DatabaseCode & "]"
    Me.Show vbModal

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "frmImportStudyDefinition.Display")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
End Sub
