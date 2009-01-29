VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRegisterDatabase 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Register database"
   ClientHeight    =   1965
   ClientLeft      =   4380
   ClientTop       =   4485
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   120
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3060
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.OptionButton optRegisterDB 
         Caption         =   "Register Oracle database"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   2415
      End
      Begin VB.OptionButton optRegisterDB 
         Caption         =   "Register SQLServer / MSDE database"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmRegisterDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999. All Rights Reserved
'   File:       frmRegisterDatabase.frm
'   Author:     Will Casey, July 1999
'   Purpose:    To enable the user to be able to register either an Access or a SQL
'               database.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'Mo 3/7/01  Major re-write of the register database code.
'           Calls to form frmSQLDBRegister replaced by calls to frmADOConnect
'           Previous revision history removed.
'           Old/Unused code removed
' ASH 8/11/2001 Uncommented frmSetHTMLFolder.HTMLPath in Routine InsertNewDatabase
'               to fix bug raised by Validation Team
' ASH 16/04/2002 Modified routine InsertNewDatabase
' REM 11/02/03 - Removed all Access code
'              - Added routine to add users in the UserRole table to the Security databases UserDtabase table
'------------------------------------------------------------------------------------'

Option Explicit

'------------------------------------------------------------------------------------'
Private Sub cmdCancel_Click()
'------------------------------------------------------------------------------------'

'------------------------------------------------------------------------------------'

    Unload Me
    
End Sub

'------------------------------------------------------------------------------------'
Private Sub cmdOK_Click()
'------------------------------------------------------------------------------------'
' choose the correct form to display depending on which option button has been clicked
' strip the last 4 characters out for the database description using replace and right
' NCJ 28 Jan 00 - Deal with "already registered" properly
'   Unload at the end of this sub
'------------------------------------------------------------------------------------'
Dim sDescription As String
Dim sPath As String
Dim sDBPassword As String

    On Error GoTo ErrHandler

'    'Perpare the Access database Open dialog
'    dlgOpenFile.Flags = 0
'    dlgOpenFile.InitDir = App.Path
'    dlgOpenFile.Filter = " All Files(*.*)|*.*" _
'        & " *.mdb|*.mdb "
'    dlgOpenFile.FilterIndex = 2
'    dlgOpenFile.DialogTitle = "MACRO Register new database"
'    dlgOpenFile.CancelError = True
     
    'Check for Registering a SQL Server database
    If optRegisterDB(1).Value = True Then
        frmADOConnect.FormUsage = "Register"
        frmADOConnect.DatabaseType = MACRODatabaseType.sqlserver
        frmADOConnect.Show vbModal
        Unload frmADOConnect
    'Check for Registering a Oracle database
    ElseIf optRegisterDB(2).Value = True Then
        frmADOConnect.FormUsage = "Register"
        frmADOConnect.DatabaseType = MACRODatabaseType.Oracle80
        frmADOConnect.Show vbModal
        Unload frmADOConnect
    'Check for Registering a Access database
'    ElseIf optRegisterDB(0).Value = True Then
'         dlgOpenFile.InitDir = goUser.Database.DatabaseLocation   'SDM 15/02/00 SR2968
'         dlgOpenFile.ShowOpen
'         sPath = dlgOpenFile.FileName
'         'Description is file name without extension
'         sDescription = dlgOpenFile.FileTitle
'         sDescription = Replace(sDescription, Right(sDescription, 4), "")
'
'        'Check to see if database name is numeric if so disallow it.
'        If IsNumeric(sDescription) Then
'            DialogError ("You may not register a number as a database name.")
'            Exit Sub
'        End If
'
'        'Check that the databasename is not greater than 15 characters
'        If Len(sDescription) > 15 Then
'            DialogError ("You cannot register a database with a name that is greater than 15 characters long.")
'            Exit Sub
'        End If
'
'        'Check that the Access database has not already been registered
'        If DoesDatabaseExist(sDescription) Then
'            DialogInformation ("This database is registered already in MACRO.")
'        Else
'            ' Database not already registered. Need to prompt for its password
'            sDBPassword = InputBox("Please enter the password for this database", "MACRO Register Database")
'            If InsertNewDatabase(sDescription, sPath, sDBPassword) Then
'                DialogInformation ("This database has been registered successfully.")
'                'Refresh the list of databases on frmDatabases
'                'frmDatabases.RefreshDatabases
'                Call frmMenu.RefereshDatabaseInfoForm
'            End If
'        End If
    End If
    Unload Me
        
Exit Sub
ErrHandler:
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




''----------------------------------------------------------------------------------------------
'Public Function InsertNewDatabase(sDescription As String, _
'                            sPath As String, _
'                            sPassword As String) As Boolean
''----------------------------------------------------------------------------------------------
'' Register new Access database in Security DB
'' Assume database NOT already registered
'' NCJ 21 Jan 00 - Added password field
''----------------------------------------------------------------------------------------------
'Dim sSQL As String
'Dim ADODBConnection As ADODB.Connection
'Dim rsTest As ADODB.Recordset
'Dim sHTMLPath As String
'Dim sSHTMLPATH As String
'
'    InsertNewDatabase = False
'
'    ' Trap password problems
'    On Error GoTo ErrCantOpenDB
'
'    'connect to the database with this password
'    ' and test that it's a Macro Database
'    Set ADODBConnection = New ADODB.Connection
'    'JET OLE DB NATIVE PROVIDER
'    ADODBConnection.Open Connection_String(CONNECTION_MSJET_OLEDB_40, sPath, , , sPassword)
'
'    ' Trap not MACRO problems
'    On Error GoTo ErrNotMACRODB
'
'    ' Test for MACRO DB by looking for ClinicalTrialID = 0
'    sSQL = "SELECT ClinicalTrialId FROM ClinicalTrial WHERE ClinicalTrialId = 0"
'    Set rsTest = New ADODB.Recordset
'    rsTest.Open sSQL, ADODBConnection, adOpenKeyset, , adCmdText
'    If rsTest.RecordCount > 0 Then
'        ' Seems OK
'        rsTest.Close
'        Set rsTest = Nothing
'        ADODBConnection.Close
'        Set ADODBConnection = Nothing
'    Else
'        ' Made up error number to force jump to ErrNotMACRODB below
'        Err.Raise vbObjectError + 9000, , "Not a MACRO Database"
'    End If
'
'    ' Normal error handler
'    On Error GoTo ErrHandler
'
'    'ASH 8/11/2001
'    'Validation team Bug on Macro 2.2.3 No.4
'    'ASH 15/04/02, added new parameter to frmSetHTMLFolder
'    If Not frmSetHTMLFolder.HTMLPath(sHTMLPath, sSHTMLPATH) Then
'        'DialogWarning "Database registration cancelled"
'        Exit Function
'    End If
'
'    'SDM 26/01/00 SR2794 Added the HTMLLocation bit
'    'ZA 22/08/01, replace HMLPath with null string
'    'Mo Morris 20/9/01 Db Audit (DatabaseDescription to DatabaseCode)
'    'Mo Morris 21/9/01, DatabaseType (which used to be 0 by default) added to SQL statement
'    'ASH 9/11/01, replace HMLPath with string string variable
'    'ASH 15/04/02, added new field SecureHTMLLocation to SQL Query
'    sSQL = " INSERT INTO Databases " _
'    & "(DataBaseCode, DatabaseType, DatabaseLocation, DatabasePassword, HTMLLocation, SecureHTMLLocation)" _
'    & " VALUES ('" & sDescription & "'," & MACRODatabaseType.Access & ",'" & sPath & "','" & sPassword & "', '" & sHTMLPath & "','" & sSHTMLPATH & "'" & ")"
'    SecurityADODBConnection.Execute sSQL, , adCmdText
'    InsertNewDatabase = True
'
'Exit Function
'
'ErrCantOpenDB:
'    Set rsTest = Nothing
'    Set ADODBConnection = Nothing
'    MsgBox "The database password is not correct", vbOKOnly, "MACRO"
'    Exit Function
'
'ErrNotMACRODB:
'    Set rsTest = Nothing
'    ADODBConnection.Close
'    Set ADODBConnection = Nothing
'    MsgBox "This is not a MACRO database", vbOKOnly, "MACRO"
'    Exit Function
'
'ErrHandler:
'  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "InsertNewDatabase")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            Call MACROEnd
'   End Select
'
'End Function

'----------------------------------------------------------------------------------------------
Private Sub Form_Load()
'----------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------

    Me.Icon = frmMenu.Icon
    FormCentre Me
    
End Sub

