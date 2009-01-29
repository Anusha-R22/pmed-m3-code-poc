VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNewDatabase 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create New Database"
   ClientHeight    =   2865
   ClientLeft      =   5565
   ClientTop       =   3840
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Server/Site"
      Height          =   1000
      Left            =   60
      TabIndex        =   6
      Top             =   1020
      Width           =   3975
      Begin VB.TextBox txtSiteName 
         Height          =   315
         Left            =   2280
         MaxLength       =   8
         TabIndex        =   9
         Top             =   550
         Width           =   1455
      End
      Begin VB.OptionButton optSite 
         Caption         =   "Site"
         Height          =   345
         Left            =   120
         TabIndex        =   8
         Top             =   550
         Width           =   700
      End
      Begin VB.OptionButton optServer 
         Caption         =   "Server"
         Height          =   345
         Left            =   120
         TabIndex        =   7
         Top             =   230
         Width           =   800
      End
      Begin VB.Label Label1 
         Caption         =   "Site Name"
         Height          =   315
         Left            =   1440
         TabIndex        =   10
         Top             =   600
         Width           =   795
      End
   End
   Begin MSComDlg.CommonDialog dlgSaveScriptAs 
      Left            =   3540
      Top             =   660
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkToFile 
      Caption         =   "Write DB creation script to file"
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   2100
      Width           =   2475
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1680
      TabIndex        =   3
      Top             =   2460
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   2880
      TabIndex        =   4
      Top             =   2460
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Caption         =   "Database"
      Height          =   1000
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.OptionButton optOracleMacro 
         Caption         =   "New Oracle MACRO Database"
         Height          =   345
         Left            =   120
         TabIndex        =   2
         Top             =   550
         Width           =   2895
      End
      Begin VB.OptionButton optSQLMacro 
         Caption         =   "New SQL Server / MSDE MACRO Database"
         Height          =   345
         Left            =   120
         TabIndex        =   1
         Top             =   230
         Width           =   3555
      End
   End
End
Attribute VB_Name = "frmNewDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999. All Rights Reserved
'   File:       frmNewDataBase.frm
'   Author:     Will Casey, September 1999
'   Purpose:    To allow the user to create a new blank database in either SQL
'               or Access for the Macro database and the security database
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'Revisions:
'Mo 3/7/01  Major re-write of the create database code.
'           Previous revision history removed.
'           Old/Unused code removed
'Mo 24/8/01 CreateDB routines moved to modNewDatabaseAndTables
'Ash 13/06/2002 Fix for outstanding 2.2.5 bug to do with creation of read only databases
'ash 12/9/2002 Commented out ACCESS code
'TA 03/11/2002: Use scripts to create new MACRO db
'REM 15/11/02 - added Server/Site settings
'------------------------------------------------------------------------------------'

Option Explicit

Private mbWriteDBScriptToFile As Boolean

Private moMacroADODBConnection As ADODB.Connection
Private mbIsLoading As Boolean



'---------------------------------------------------------------------
Private Sub chkToFile_Click()
'---------------------------------------------------------------------

    If chkToFile.Value = vbChecked Then
        mbWriteDBScriptToFile = True
    Else
        mbWriteDBScriptToFile = False
    End If

End Sub

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------
 
    On Error GoTo ErrHandler
    
    mbIsLoading = True
    
    'set the "Write DB creation script to file" checkbox to unchecked
    'as well as the boolean flag mbWriteDBScriptToFile
    mbWriteDBScriptToFile = False
    chkToFile.Value = vbUnchecked
    optServer.Value = True
    Me.Icon = frmMenu.Icon
    FormCentre Me
    mbIsLoading = False

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

'----------------------------------------------------------------------------------------------
Private Sub cmdCancel_Click()
'----------------------------------------------------------------------------------------------
    
    Unload Me
    
End Sub

'---------------------------------------------------------------------------------------------
Public Sub cmdOK_Click()
'---------------------------------------------------------------------------------------------
' When creating Macro Access databases the user has to decide the name and where to save the database.
' When writing the DB creation script to file the user has to decide the name/location of the script file.
' When creating Macro SQL Server and Oracle databases the DB into which the Macro tables will be
' created must already exist on the server, hence the User is not involved in deciding where its saved to.
'---------------------------------------------------------------------------------------------

Dim sPath As String
Dim sDescription As String
Dim sSQLServer As String
Dim sSQLDatabase As String
Dim sConnectionStringOrPath As String
Dim sMSG As String
'Ash 13/06/2002
Dim objFSO As New FileSystemObject
Dim objFile As File
Dim bFileExists As Boolean
Dim sSiteSettingValue As String
Dim sCon As String
Dim sMACROFileName As String
Dim sSiteCode As String

    On Error GoTo ErrorHandler
    
    'REM 15/11/02 - added Server/Site settings
    If optSite.Value Then
        sSiteSettingValue = txtSiteName.Text
    Else
        sSiteSettingValue = ""
    End If
    
    'check for creating a Macro database (SQL Server)
    '------------------------------------------------
    If optSQLMacro.Value = True Then
        'Check for ToScript or an actual build
        If mbWriteDBScriptToFile Then
            sMACROFileName = InputBox("Please enter the name of the MACRO database script file", "MACRO Script File", "MACRODB")
            If sMACROFileName = "" Then Exit Sub
            sSiteCode = InputBox("Please enter a site name if this is a site database", "MACRO Script File")
            Call CreateNewMACRODB(frmADOConnect.ConnectString, MACRODatabaseType.sqlserver, True, sMACROFileName, , sSiteCode)
            
        Else
            frmADOConnect.FormUsage = "Create"
            frmADOConnect.DatabaseType = MACRODatabaseType.sqlserver
            frmADOConnect.Show vbModal
            sCon = frmADOConnect.ConnectString
            If sCon <> vbNullString Then
                
                Call CreateNewMACRODB(frmADOConnect.ConnectString, MACRODatabaseType.sqlserver, False, "", , sSiteSettingValue)
               
               'see if user wants to register database
                Call RegisterNewMACRODB(sCon, MACRODatabaseType.sqlserver)
               
            End If
            Unload frmADOConnect
        End If
        
    'check for creating a Macro database (Oracle)
    '--------------------------------------------
    ElseIf optOracleMacro.Value = True Then
        'Check for ToScript or an actual build
        If mbWriteDBScriptToFile Then
            sMACROFileName = InputBox("Please enter the name of the MACRO database script file", "MACRO Script File", "MACRODB")
            If sMACROFileName = "" Then Exit Sub
            sSiteCode = InputBox("Please enter a site name if this is a site database", "MACRO Script File")
            Call CreateNewMACRODB(frmADOConnect.ConnectString, MACRODatabaseType.oracle80, True, sMACROFileName, , sSiteCode)
            
        Else
            frmADOConnect.FormUsage = "Create"
            frmADOConnect.DatabaseType = MACRODatabaseType.oracle80
            frmADOConnect.Show vbModal
            sCon = frmADOConnect.ConnectString
            If sCon <> vbNullString Then
                Call CreateNewMACRODB(frmADOConnect.ConnectString, MACRODatabaseType.oracle80, False, , , sSiteSettingValue)
                
                Call RegisterNewMACRODB(sCon, MACRODatabaseType.oracle80)
                
            End If
            Unload frmADOConnect
        End If
    End If
  
    Unload Me
   
Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 32755
            'Error number 32755 means that a common dialog box has been canceled
            Exit Sub
        Case Else
            Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdOK_Click")
                Case OnErrorAction.Ignore
                    Resume Next
                Case OnErrorAction.Retry
                    Resume
                Case OnErrorAction.QuitMACRO
                    Call ExitMACRO
                    Call MACROEnd
            End Select
    End Select
        
End Sub

'---------------------------------------------------------------------------------------------
Private Sub RegisterNewMACRODB(sCon As String, nDatabaseType As Integer)
'---------------------------------------------------------------------------------------------
'REM 01/04/03
'Register MACRO database
'---------------------------------------------------------------------------------------------
Dim sServer As String
Dim sDatabase As String
Dim sDBAlias As String
Dim sDBPswd As String
Dim sDBUser As String
Dim tCon As udtConnection
Dim sDataSource As String

     If DialogQuestion("Would you like to register this database now?", "Register", False) = vbYes Then
     
        tCon = Connection_AsType(sCon)

        sServer = tCon.Datasource
        sDBUser = tCon.UserId
        sDatabase = tCon.Database
        sDBPswd = tCon.Password
        
        If nDatabaseType = MACRODatabaseType.oracle80 Then
            sDatabase = sServer
            sServer = ""
        End If
        
        'Get the new database alias, routine check's to see if it is unique
        sDBAlias = NewDBAlias(sDatabase)
        
        If sDBAlias = "" Then
            Exit Sub
        End If

        'register the newly created database
        Call RegisterMACRODatabase(sCon, sDBAlias, nDatabaseType, sServer, sDatabase, sDBPswd, sDBUser)

     Else
         DialogInformation "MACRO database created." & vbCrLf _
             & "You must register the database before you can use it for data entry."
     End If

End Sub


'--------------------------------------------------------------------
Private Sub optServer_Click()
'--------------------------------------------------------------------
    
    cmdOK.Enabled = True
    txtSiteName.Enabled = False
End Sub

'--------------------------------------------------------------------
Private Sub optSite_Click()
'--------------------------------------------------------------------
'REM 15/11/02
'When user clicks teh Site option button must disable the OK button if there is no site name in the Site Name text box
'--------------------------------------------------------------------
    
    txtSiteName.Enabled = True
    
    If txtSiteName.Text = "" Then
        cmdOK.Enabled = False
    End If

End Sub

'--------------------------------------------------------------------
Private Sub txtSiteName_Change()
'--------------------------------------------------------------------
'REM 15/11/02
'Only enable the OK button if there is text in the text box
'--------------------------------------------------------------------
Dim sCode As String
Dim iPos As Integer

    iPos = txtSiteName.SelStart

    cmdOK.Enabled = (txtSiteName <> "")

    'If the form is loading then don't run the change event procedures
    If mbIsLoading = False Then
        'make the site name lower case only
        txtSiteName.Text = LCase(txtSiteName)
        
        sCode = txtSiteName.Text
        ' check that the character entered is valid
        If Not gblnValidString(sCode, valAlpha + valNumeric) Then
            txtSiteName.Text = txtSiteName.Tag
            txtSiteName.SelStart = 0
            txtSiteName.SelLength = Len(txtSiteName.Text)
            
        End If
    
        ' check that the first char is not numeric
        If sCode <> vbNullString Then
            If gblnValidString(Left$(sCode, 1), valNumeric) Then
            txtSiteName.Text = txtSiteName.Tag
            txtSiteName.SelStart = 0
            txtSiteName.SelLength = Len(txtSiteName.Text)
                
            End If
        End If

    End If

    If iPos > 0 Then
        txtSiteName.SelStart = iPos
    End If

    txtSiteName.Tag = Trim(txtSiteName.Text)

End Sub

