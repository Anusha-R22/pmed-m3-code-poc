VERSION 5.00
Begin VB.Form frmExportLab 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set at run time"
   ClientHeight    =   1290
   ClientLeft      =   8055
   ClientTop       =   9735
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "set at runtime"
      Height          =   675
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   3975
      Begin VB.ComboBox cboLabs 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   3765
      End
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "cmsStart"
      Height          =   375
      Left            =   2820
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmExportLab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   File:       frmExportLab.frm
'   Author:     Mo Morris October 2000
'   Purpose:    Used for selecting and distributing Laboratory Definitions.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'TA 19/10/2000: Distribute now user this form aswell
'DPH 17-18/10/2001 Added FolderExistence routine calls to create missing folders
'MLM 23/04/07: Bug 1522: Added notification of successful export.
'----------------------------------------------------------------------------------------'
Option Explicit
Option Compare Binary
Option Base 0

'true if distributing
'false if exporting
Private mbDistribute As Boolean

Private msSelectedLab As String
'ASH 13/12/2002
Private msDatabase As String
Private oDatabase As MACROUserBS30.Database
Private bLoad As Boolean
Private sConnectionString As String
Private sMessage As String
Private mconMACRO As ADODB.Connection


'---------------------------------------------------------------------
Private Sub DistributeLab()
'---------------------------------------------------------------------
' REVISIONS
' DPH 08/05/2002 - Don't write message to message table unless
'           file is placed in published HTML folder
'---------------------------------------------------------------------
Dim oExchange As clsExchange
Dim sCabFileName As String
Dim rsLabSites As ADODB.Recordset
Dim sSQL As String
Dim lMessageId As Long

    On Error GoTo ErrHandler
    
    Screen.MousePointer = vbHourglass
    
    Set oExchange = New clsExchange
    
    'export the selected laboratory into sCabFileName
    sCabFileName = oExchange.ExportLDD(msSelectedLab)
    
    ' DPH 17/10/2001 - Check cab file returns filename
    If sCabFileName <> "" Then
        ' Do not write message to message table unless successfully copied

        'copy the lab cab file to the trialoffice html folder
        Do Until FileExists(gsOUT_FOLDER_LOCATION & sCabFileName)
            DoEvents
        Loop
        
        On Error GoTo CopyFileErr
        
        ' DPH 17/10/2001 Make sure folder exists before opening
        ' DPH 16/04/2002 - Corrected (as was checking for file existence)
        If FolderExistence(goUser.Database.HTMLLocation & sCabFileName, False) Then
            Call FileCopy(gsOUT_FOLDER_LOCATION & sCabFileName, _
                            goUser.Database.HTMLLocation & sCabFileName)
        End If
        
        On Error GoTo ErrHandler
        
        sSQL = "SELECT Site FROM SiteLaboratory WHERE LaboratoryCode = '" & msSelectedLab & "'"
        Set rsLabSites = New ADODB.Recordset
        rsLabSites.Open sSQL, mconMACRO, adOpenForwardOnly, adOpenStatic, adCmdText
        
        Do While Not rsLabSites.EOF
            'Mo Morris 30/8/01 Db Audit (UserId to UserName, MessageId no longer an autonumber)
            lMessageId = NextMessageId(mconMACRO)
            sSQL = "INSERT INTO Message (MessageId, TrialSite, MessageType, MessageTimestamp, UserName, " _
                & "MessageReceived, MessageDirection, MessageBody, MessageParameters) " _
                & "  VALUES (" & lMessageId & ",'" & rsLabSites![Site] & "'," & ExchangeMessageType.LabDefinitionServerToSite & "," _
                & SQLStandardNow & ",'" & goUser.UserName & "'," & MessageReceived.NotYetReceived & "," _
                & MessageDirection.MessageOut & ",'Laboratory " & msSelectedLab & " is being distributed to your site." _
                & "','" & sCabFileName & "')"
            mconMACRO.Execute sSQL
                
            rsLabSites.MoveNext
        Loop
        
    Else
        Screen.MousePointer = vbDefault
        Call DialogError("Export failed as could not create file", "Lab Definition Export")
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    
    Call DialogInformation(msSelectedLab & " has successfully been distributed", "Lab Definition Export")
    
    cmdStart.Enabled = False

Exit Sub
CopyFileErr:
    ' If an error copying file
    Screen.MousePointer = vbDefault
    
    Call DialogError("Distribute laboratory aborted - Error copying file to published HTML folder", "Lab Definition Export")

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "DistributeLab")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
End Sub


'---------------------------------------------------------------------
Private Sub ExportLab()
'---------------------------------------------------------------------
Dim oExchange As clsExchange
Dim sCabFileName As String

    'On Error GoTo ErrHandler
    
    Screen.MousePointer = vbHourglass
    
    Set oExchange = New clsExchange
    
    sCabFileName = oExchange.ExportLDD(msSelectedLab)

    Screen.MousePointer = vbDefault
    'MLM 23/04/07: Bug 1522:
    Call DialogInformation(msSelectedLab & " has successfully been exported", "Laboratory Definition Export")

    cmdStart.Enabled = False

End Sub

'---------------------------------------------------------------------
Private Sub cmdStart_Click()
'---------------------------------------------------------------------
    
    If mbDistribute Then
        Call DistributeLab
    Else
        Call ExportLab
    End If
    
End Sub

'---------------------------------------------------------------------
Public Sub Display(bDistribute As Boolean, sDatabase As String)
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsLabs As ADODB.Recordset

    On Error GoTo ErrHandler
    
    msDatabase = sDatabase
    Set oDatabase = New MACROUserBS30.Database
    bLoad = oDatabase.Load(SecurityADODBConnection, goUser.UserName, msDatabase, "", False, sMessage)
    sConnectionString = oDatabase.ConnectionString
    Set mconMACRO = New ADODB.Connection
    mconMACRO.Open sConnectionString
    mconMACRO.CursorLocation = adUseClient

    
    mbDistribute = bDistribute
    
    'set form appearance according to distribute/export
    If mbDistribute Then
        Me.Caption = "Distribute Laboratory " & "[" & goUser.DatabaseCode & "]"
        Frame1.Caption = "Select laboratory to be distributed"
        cmdStart.Caption = "&Distribute"
    Else
        Me.Caption = "Export Laboratory " & "[" & goUser.DatabaseCode & "]"
        Frame1.Caption = "Select laboratory to be exported"
        cmdStart.Caption = "&Export"
    End If
        
    FormCentre Me
    cmdStart.Enabled = False
    
    sSQL = "SELECT LaboratoryCode FROM Laboratory ORDER BY LaboratoryCode"
    Set rsLabs = New ADODB.Recordset
    rsLabs.Open sSQL, mconMACRO, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    cboLabs.Clear
    Do Until rsLabs.EOF
        cboLabs.AddItem rsLabs!LaboratoryCode
        rsLabs.MoveNext
    Loop
    
    rsLabs.Close
    Set rsLabs = Nothing
    
    Me.Show vbModal

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "Display")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select


End Sub

'---------------------------------------------------------------------
Private Sub cmdOK_Click()
'---------------------------------------------------------------------

    Unload Me
 
End Sub

'---------------------------------------------------------------------
Private Sub cboLabs_click()
'---------------------------------------------------------------------
On Error GoTo ErrHandler

    If cboLabs.ListIndex = -1 Then
        cmdStart.Enabled = False
        Exit Sub
    Else
        cmdStart.Enabled = True
    End If
    
    msSelectedLab = cboLabs.Text

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cboLabs_click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
   
End Sub

'----------------------------------------------------------------
Private Sub Form_Load()
'----------------------------------------------------------------

    Me.Icon = frmMenu.Icon
    FormCentre Me

End Sub
