VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportPatientData 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Subjects"
   ClientHeight    =   975
   ClientLeft      =   570
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
   Begin VB.TextBox txtImportMessage 
      Enabled         =   0   'False
      Height          =   372
      Left            =   4800
      TabIndex        =   4
      Top             =   480
      Width           =   3795
   End
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
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3012
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3885
      Top             =   0
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Import Progress:"
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   210
      Visible         =   0   'False
      Width           =   1215
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
Attribute VB_Name = "frmImportPatientData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       frmImportPatientData.frm
'   Author:     Mo Morris, January 1998
'   Purpose:    Allows selection of file containing patient data to be imported.
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
'   Revisions:
'   1    Mo Morris          26/01/98
'   2    Mo Morris          24/02/98
'   3    Andrew Newbigging  2/04/98
'   4    Joanne Lau         13/05/98
'   5    Joanne Lau         18/06/98
'   6    Joanne Lau         29/06/98
'   7    Andrew Newbigging  20/11/98
'       Modified to do COMPACT import as well as standard MACRO import.
'       ImportType property is used to indicate which.
'       Mo Morris           18/2/99
'       Changes stemming from COMPACT import now handled by frmImportCompactPatientData
'       and the fact that standard patient import is from a CAB file not a prd file
'   8    PN                 14/09/99 Upgrade to VB standards doc version 1.0
'                           Set window border to fixed dialog
'   PN  21/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   WillC 10/11/99  Added the error handlers
'   Mo Morris   2/3/00 cmdStartImport_Click now imports multimedia files via a call to ImportDocumentsAndGraphics
'   TA 08/05/2000   subclassing removed
'   TA 19/10/2000: Use new hourglass functions to ensure hourglass displayed
'   DPH 17/10/2001 Added FolderExistence routine calls to create missing folders
'   DPH 18/04/2002 - Include ZIP files
'   RS  27/02/2003 - Report 'Trial does not exist' error back to user
'--------------------------------------------------------------------------------
    
'---------------------------------------------------------------------
Option Explicit
Option Compare Binary
Option Base 0
'---------------------------------------------------------------------
Private gsMACRO_IMPORTS As String
Private gsImportFile As String
'ASH 11/12/2002
Private oDatabase As MACROUserBS30.Database
Private bLoad As Boolean
Private sConnectionString As String
Private sMessage As String
Private mconMACRO As ADODB.Connection
Private msDatabase As String
Private mbLabData As Boolean

'---------------------------------------------------------------------
Private Sub cmdSelectImportFile_Click()
'---------------------------------------------------------------------
' REVISIONS
' DPH 18/04/2002 - Include ZIP files
'---------------------------------------------------------------------

    cmdStartImport.Enabled = False
    CommonDialog1.CancelError = True
    
    On Error GoTo ErrInShowOpen
    CommonDialog1.ShowOpen
    gsImportFile = CommonDialog1.FileName
    
    'Mo 24/1/00, file extention check added
    If LCase(Mid(gsImportFile, Len(gsImportFile) - 3, 4)) <> ".cab" And _
        LCase(Mid(gsImportFile, Len(gsImportFile) - 3, 4)) <> ".zip" Then
        MsgBox ("Patient data Import files must have a 'cab' or 'zip' extension")
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
'Mo Morris 28/4/00, SR3249, changes made so that the routine exits properly
'when an error occurs during an export.
'---------------------------------------------------------------------

Dim msNextPRDFile As String
Dim mExchange As clsExchange
Dim lResult As ExchangeError     ' Store Result of ImportPRD

Set mExchange = New clsExchange

    On Error GoTo ErrHandler

    HourglassOn
    
    If Not mbLabData Then
    
        'Unpack the CAB file into .prd and .psf files and place them
        'in directory AppPath/CabExtract
        txtImportMessage.Text = "Extracting " & gsImportFile
        DoEvents
        Call mExchange.ImportPatientCAB(gsImportFile)
        
        'loop through the extracted files and import them into Macro
        
        'SDM 26/01/00 SR2794
        msNextPRDFile = Dir(gsCAB_EXTRACT_LOCATION & "*.prd")
    '    msNextPRDFile = Dir(gsAppPath & "CabExtract" & "\*.prd")
        Do While msNextPRDFile <> ""
            txtImportMessage.Text = "Importing " & msNextPRDFile
            DoEvents
            'SDM 26/01/00 SR2794
            lResult = mExchange.ImportPRD(gsCAB_EXTRACT_LOCATION & msNextPRDFile)
            Select Case lResult
    '        Select Case mExchange.ImportPRD(gsAppPath & "CabExtract" & "\" & msNextPRDFile)
            Case ExchangeError.EmptyFile
                MsgBox (gsImportFile & " is empty." + vbNewLine + "Import aborted.")
                HourglassOff
                cmdStartImport.Enabled = False
                Exit Sub
            Case ExchangeError.Invalid
                MsgBox (gsImportFile & " is not a valid patient response file." + vbNewLine + "Import aborted.")
                HourglassOff
                cmdStartImport.Enabled = False
                Exit Sub
            ' DPH 17/10/2001
            Case ExchangeError.DirectoryNotFound
                MsgBox (gsImportFile & " could not be found." + vbNewLine + "Import aborted.")
                HourglassOff
                cmdStartImport.Enabled = False
                Exit Sub
            Case ExchangeError.Success
                'MsgBox ("Import successfully completed.")
                txtImportMessage.Text = msNextPRDFile & " imported."
            ' DPH 10/05/2002 - Trial Locked Error
            Case ExchangeError.TrialLocked
                MsgBox ("Cannot obtain a study lock." + vbNewLine + "Import aborted.")
                HourglassOff
                cmdStartImport.Enabled = False
                Exit Sub
            ' RS 27/02/2003 - Error was reported by ImportPRD, but not handled separately
            Case ExchangeError.TrialDoesntExist
                MsgBox ("Trial does not exist." + vbNewLine + "Import aborted.")
                HourglassOff
                cmdStartImport.Enabled = False
                Exit Sub
            'Mo 21/11/2005 COD0220
            Case ExchangeError.MoreRecentVersion
                MsgBox ("Import of Patient Data from a more recent version of MACRO disallowed." + vbNewLine + "Import aborted.")
                HourglassOff
                cmdStartImport.Enabled = False
                Exit Sub
            Case Else
                MsgBox ("Unexpected error. Import aborted")
                HourglassOff
                cmdStartImport.Enabled = False
                Exit Sub
            End Select
            'get next prd file via the DIR command
            msNextPRDFile = Dir
        Loop
        
        ' DPH 17/10/2001 Make sure folder exists before opening
        If FolderExistence(gsImportFile, True) Then
            Kill gsImportFile
            'changed Mo Moris 2/3/00
            Do Until Not FileExists(gsImportFile)
                DoEvents
            Loop
            txtImportMessage.Text = "Import completed."
        Else
            txtImportMessage.Text = "Import Failed - could not find file."
        End If
        
        'changed Mo Moris 2/3/00
        'move any patient data multimedia files from the CabExtract folder to the Documents folder
        mExchange.ImportDocumentsAndGraphics
    Else
        ImportLab
        End If
    
    HourglassOff
    cmdStartImport.Enabled = False

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
' REVISIONS
' DPH 18/04/2002 - Include ZIP files
'---------------------------------------------------------------------
     On Error GoTo ErrHandler
 
    Me.Icon = frmMenu.Icon
        
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

'-------------------------------------------------------------------------------
Public Sub Display(ByVal sDatabase As String, Optional bLab As Boolean = False)
'-------------------------------------------------------------------------------
'
'-------------------------------------------------------------------------------
    
    msDatabase = sDatabase
    mbLabData = bLab
    Set oDatabase = New MACROUserBS30.Database
    bLoad = oDatabase.Load(SecurityADODBConnection, goUser.UserName, msDatabase, "", False, sMessage)
    sConnectionString = oDatabase.ConnectionString
    Set mconMACRO = New ADODB.Connection
    mconMACRO.Open sConnectionString
    mconMACRO.CursorLocation = adUseClient
    
    gsMACRO_IMPORTS = gsIN_FOLDER_LOCATION
    
    With CommonDialog1
        .DialogTitle = "Import File Selection"
        .InitDir = gsMACRO_IMPORTS
        .DefaultExt = "cab"
        'Mo 24/1/00
        '.Filter = "Patient Data (*.cab;*.1;*.*)|*.cab;*.1;*.*"
        .Filter = "Patient Data (*.cab;*.zip)|*.cab;*.zip"
    End With
    
    If bLab Then
        Me.Caption = "Import Laboratory Data " & "[" & goUser.DatabaseCode & "]"
    Else
       Me.Caption = "Import Subjects " & "[" & goUser.DatabaseCode & "]"
    End If
    
    cmdStartImport.Enabled = False
    Me.Show vbModal
End Sub

'-------------------------------------------------------------------------------------
Private Sub ImportLab()
'-------------------------------------------------------------------------------------
'since no seperate form exists in EX for lab imports, this routine will be used instead
'-------------------------------------------------------------------------------------
Dim sNextLDDFile As String
Dim oExchange As clsExchange
Dim sImportFile As String

    On Error GoTo ErrHandler
  
    'sImportFile = gsIN_FOLDER_LOCATION & "*.*"
    'If CMDialogOpen(CommonDialog1, "Select Laboratory Definition Import File", sImportFile, "Laboratory Definition Import Files (*.cab;*.zip)|*.cab;*.zip") Then
  
        If DialogQuestion("Are you sure you wish to import laboratory definition import file" & vbCrLf & gsImportFile) = vbYes Then
            HourglassOn
            
            Set oExchange = New clsExchange
        
            'Unpack the CAB file into an .ldd file into directory AppPath/CabExtract
            oExchange.ImportLDDCAB (gsImportFile)
            
            'loop through the extracted files and import them into Macro
            sNextLDDFile = Dir(gsCAB_EXTRACT_LOCATION & "*.ldd")
            If sNextLDDFile = "" Then
                'no ldd files in cab
                Call MsgBox("No laboratory definition file found" + vbNewLine + "Import aborted.", , "Import Laboratory")
            Else
                'display status form
                'Call frmStatus.Start(GetApplicationTitle, "Importing laboratory definition " & StripFileNameFromPath(sNextLDDFile) & "...", False)
                
                Select Case oExchange.ImportLDD(gsCAB_EXTRACT_LOCATION & sNextLDDFile)
                Case ExchangeError.EmptyFile
                    'Call frmStatus.Finish
                    Call MsgBox(sImportFile & " is empty." + vbNewLine + "Import aborted.", , "Import Laboratory")
                Case ExchangeError.Invalid
                    'Call frmStatus.Finish
                    Call MsgBox(sImportFile & " is not a valid laboratory definition file." + vbNewLine + "Import aborted.", , "Import Laboratory")
                Case ExchangeError.DirectoryNotFound
                    'Call frmStatus.Finish
                    Call MsgBox(sImportFile & " does not exist." + vbNewLine + "Import aborted.", , "Import Laboratory")
                Case ExchangeError.Success
                    'Call frmStatus.Finish
                    Call MsgBox(sNextLDDFile & " imported.", , "Import Laboratory")
                Case Else
                    'Call frmStatus.Finish
                    Call MsgBox("Unexpected error. Import aborted", , "Import Laboratory")
                End Select
            End If
            
            HourglassOff
        End If
    
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "cmdImportLab_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub
