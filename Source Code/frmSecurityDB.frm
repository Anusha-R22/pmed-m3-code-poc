VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSecurityDB 
   Caption         =   "Security Database"
   ClientHeight    =   3750
   ClientLeft      =   9000
   ClientTop       =   4755
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   5580
   Begin MSComDlg.CommonDialog dlgBrowse 
      Left            =   3780
      Top             =   3300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      Height          =   375
      Left            =   60
      TabIndex        =   9
      ToolTipText     =   "Create a new security database"
      Top             =   3300
      Width           =   1155
   End
   Begin VB.CommandButton cmdUpgrade 
      Caption         =   "Upgrade"
      Height          =   375
      Left            =   2580
      TabIndex        =   8
      ToolTipText     =   "Upgrade an existing security database"
      Top             =   3300
      Width           =   1155
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Next Session"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      ToolTipText     =   "Set security database for the next session"
      Top             =   3300
      Width           =   1155
   End
   Begin VB.Frame fraCurrent 
      Caption         =   "Security Database for current MACRO session"
      Height          =   1545
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5460
      Begin VB.TextBox txtCurrent 
         Height          =   1035
         Left            =   1140
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   300
         Width           =   4185
      End
      Begin VB.Label Label1 
         Caption         =   "Connection parameters"
         Height          =   465
         Left            =   180
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   3300
      Width           =   1215
   End
   Begin VB.Frame fraNext 
      Caption         =   "Security Database for next MACRO session"
      Height          =   1545
      Left            =   60
      TabIndex        =   1
      Top             =   1680
      Width           =   5460
      Begin VB.TextBox txtNext 
         Height          =   1035
         Left            =   1080
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   300
         Width           =   4245
      End
      Begin VB.Label Label2 
         Caption         =   "Connection parameters"
         Height          =   405
         Left            =   180
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmSecurityDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2002-2004. All Rights Reserved
'   File:       frmSecurityDB.frm
'   Author:     Richard Meinesz, InferMed, Nov 2002
'   Purpose:    Security Database form
'
'----------------------------------------------------------------------------------------'
' REVISIONS:
'   NCJ 15 Jan 04 - Added file header; corrected "Secuirty" typos (SR5358)
'----------------------------------------------------------------------------------------'


Option Explicit

'---------------------------------------------------------------------
Public Sub Display()
'---------------------------------------------------------------------
Dim sText As String

    Me.Icon = frmMenu.Icon
    
    sText = ProviderInfo(SecurityDatabasePath)
    
    txtCurrent.Text = sText
    
    FormCentre Me
    
    Me.Show vbModal

End Sub

'---------------------------------------------------------------------
Private Sub cmdCancel_Click()
'---------------------------------------------------------------------

    Unload Me

End Sub

'---------------------------------------------------------------------
Private Sub cmdChange_Click()
'---------------------------------------------------------------------
Dim sCon As String
Dim sText As String

    fraNext.Caption = "security for next session"
    
    sCon = CreateOrRegisterSecurityDB(True, True)
    sText = ProviderInfo(sCon)
    txtNext.Text = sText
 
End Sub

'---------------------------------------------------------------------
Private Sub cmdUpgrade_Click()
'---------------------------------------------------------------------
'REM 06/06/03
'Upgrade a pre-MACRO 3.0 security database.  Give user option to either create a new security database or
'upgrade the security information into the existing database.
'---------------------------------------------------------------------
Dim lCreateImport As Long
Dim sOldSecPath As String
Dim sNewSecCon As String
Dim sSecVersion As String
    
    On Error GoTo ErrLabel

    'Prepare the Access database open dialog
    With dlgBrowse
        .Flags = 0
        .InitDir = App.Path
        .Filter = "*.mdb|*.mdb"
        .DialogTitle = "MACRO Security Database"
        .CancelError = False
    End With
    
    dlgBrowse.ShowOpen
    sOldSecPath = dlgBrowse.FileName

    'if user cancelled then returns empty string
    If sOldSecPath <> "" Then
            
        If IsValidSecurityDB(sOldSecPath, sSecVersion) Then

            lCreateImport = frmOptionMsgBox.Display(GetApplicationTitle, "Upgrade security database version " & sSecVersion, "Please select one of the following:", "Create new security database|Import security database", "&OK", "&Cancel", True, True)
            
            Select Case lCreateImport
            Case 1
                'Create a new security database and import all security info
                sNewSecCon = CreateOrRegisterSecurityDB(True, False)
                
                If sNewSecCon <> "" Then
                    'Import data from old security database
                    Call ImportSecurityData(sOldSecPath, sNewSecCon)
                    Call DialogInformation("Security database upgrade successful.")
                End If
            Case 2
                'Import security info into current security database
                Call ImportSecurityData(sOldSecPath, SecurityDatabasePath)
                Call DialogInformation("Security database upgrade successful.")
            End Select
        Else
            Call DialogError("This is not a valid MACRO security database!")
        End If
    End If
    
Exit Sub
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "frmSecurityDB.cmdUpgrade_Click"
End Sub

'---------------------------------------------------------------------
Private Function IsValidSecurityDB(sSecPath As String, ByRef sSecVersion As String) As Boolean
'---------------------------------------------------------------------
'REM 23/06/03
'Check that selected Access database is a MACRO security database
'---------------------------------------------------------------------
Dim oAccessCon As ADODB.Connection
Dim sSQL As String
Dim rsDb As ADODB.Recordset

    On Error GoTo ErrLabel

      'create connection to old Access security database
    Set oAccessCon = New ADODB.Connection
    oAccessCon.Open Connection_String(CONNECTION_MSJET_OLEDB_40, sSecPath, , , gsSecurityDatabasePassword)
    
    sSQL = "SELECT * FROM SecurityControl"
    Set rsDb = New ADODB.Recordset
    rsDb.Open sSQL, oAccessCon, adOpenKeyset, adLockPessimistic, adCmdText
    
    sSecVersion = rsDb!MACROVersion & "." & rsDb!BuildSubVersion
    
    IsValidSecurityDB = True
    
Exit Function
ErrLabel:
    IsValidSecurityDB = False
    sSecVersion = ""
End Function

'---------------------------------------------------------------------
Private Sub cmdCreate_Click()
'---------------------------------------------------------------------
Dim sCon As String
Dim sText As String

    fraNext.Caption = "security db just created"
    
    sCon = CreateOrRegisterSecurityDB(True, False)
    
    sText = ProviderInfo(sCon)
    txtNext.Text = sText
   
End Sub

'---------------------------------------------------------------------
Private Function ProviderInfo(sCon As String) As String
'---------------------------------------------------------------------
Dim sProvider As String
Dim sDataSource As String
Dim sUserId As String
Dim sDatabase As String
Dim tCon As udtConnection
Dim sText As String

    If sCon = "" Then Exit Function

    tCon = Connection_AsType(sCon)
    
    sProvider = tCon.Provider
    sDataSource = tCon.Datasource
    sUserId = tCon.UserId
    sDatabase = tCon.Database

    If sProvider = CONNECTION_MSDAORA Then
        sText = "Provider = " & sProvider & vbCrLf & "Data Source = " & sDataSource & vbCrLf & "User ID = " & sUserId
    ElseIf sProvider = CONNECTION_SQLOLEDB Then
        sText = "Provider = " & sProvider & vbCrLf & "Database = " & sDatabase & vbCrLf & "User ID = " & sUserId
    Else
        sText = "Unknown Provider: " & sProvider
    End If
    
    ProviderInfo = sText
    
End Function
