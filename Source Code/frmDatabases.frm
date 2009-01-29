VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDatabases 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Manager"
   ClientHeight    =   5880
   ClientLeft      =   5730
   ClientTop       =   4425
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   8235
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Security database for current Macro session"
      Height          =   855
      Left            =   240
      TabIndex        =   14
      Top             =   2760
      Width           =   7695
      Begin VB.TextBox txtCurrentSecurityDatabaseLocation 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2040
         TabIndex        =   15
         Top             =   360
         Width           =   5535
      End
      Begin VB.Label Label2 
         Caption         =   "Location:"
         Height          =   195
         Left            =   1320
         TabIndex        =   16
         Top             =   360
         Width           =   660
      End
   End
   Begin VB.Frame Frame3 
      Height          =   700
      Left            =   240
      TabIndex        =   8
      Top             =   5040
      Width           =   7695
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   375
         Left            =   4920
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Close"
         Height          =   375
         Left            =   6360
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Security database for next Macro session"
      Height          =   1250
      Left            =   240
      TabIndex        =   2
      Top             =   3720
      Width           =   7695
      Begin VB.CommandButton cmdDefaultSecurity 
         Caption         =   "&Default"
         Height          =   375
         Left            =   6360
         TabIndex        =   7
         Top             =   765
         Width           =   1215
      End
      Begin VB.CommandButton cmdBrowser 
         Caption         =   "&Browse"
         Height          =   375
         Left            =   4920
         TabIndex        =   6
         Top             =   765
         Width           =   1215
      End
      Begin VB.TextBox txtSecurityDatabaseLocation 
         Height          =   300
         Left            =   2040
         TabIndex        =   4
         Top             =   360
         Width           =   5535
      End
      Begin VB.CheckBox chkProtectSecurityDatabase 
         Caption         =   "Password protected"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   765
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Location:"
         Height          =   195
         Left            =   1320
         TabIndex        =   5
         Top             =   435
         Width           =   660
      End
   End
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   7560
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraActions 
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   7695
      Begin VB.CommandButton cmdDBPassword 
         Caption         =   "Set database password"
         Height          =   495
         Left            =   5280
         TabIndex        =   13
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton cmdRegisterDb 
         Caption         =   "Register database..."
         Height          =   495
         Left            =   2760
         TabIndex        =   12
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton cmdCreateDB 
         Caption         =   "Create new database..."
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   2295
      End
   End
   Begin MSComctlLib.ListView lvwDatabases 
      Height          =   1455
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmDatabases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999. All Rights Reserved
'   File:       frmDatabases.frm
'   Author:     Will Casey, 12/10/99
'   Purpose:    Allows the user to create a new database or edit an existing database
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'  willC   9/12/99  Added the CheckUserDatabaseRights sub
'  WillC  20/1/00   Added check to make sure the default database we choose
'                   is a security database
'   NCJ 21/1/00 Moved InsertNewDatabase to frmRegisterDatabases
'   NCJ 24/1/00 SR1225 Tidied up registering of security databases
'   NCJ 25/1/00 Bug fixing yesterday's work
'   NCJ 28/1/00 SR1225 Deal correctly with Cancel in ChangeDatabasePassword dialog
'   WillC 8/5/00 SR3336 Allow users to change the password for non Access database in cmdDatabases.
'   ATO 10/06/2002 Fix for bug 2.2.8 no.8
'------------------------------------------------------------------------------------'

' NCJ 24/1/00 - Store whether it's a programmatic check box click
Private mbProgramChecking As Boolean
'ATO 10/06/2002
Private msNextSecurityDatabase As String

Option Explicit

'----------------------------------------------------------------------------------------'
Private Sub chkProtectSecurityDatabase_Click()
'----------------------------------------------------------------------------------------'

Dim oWorkspace As dao.Workspace
Dim oDatabase As dao.Database
Dim blnChanged As Boolean

    On Error GoTo ErrHandler

    ' NCJ 24/1/00 - Check for programmatic update
    If mbProgramChecking Then
        ' Reset for user action
        mbProgramChecking = False
        Exit Sub
    End If
    
    If chkProtectSecurityDatabase.Value = 2 Then
        ' It's greyed out so ignore
        Exit Sub
    End If
    
    ' Temporarily shut down our own security DB connection
    TerminateSecurityADODBConnection
    
    '   Open the database in exclusive mode
    Set oWorkspace = DBEngine.Workspaces(0)
    
    blnChanged = False
    
    On Error Resume Next

    Set oDatabase = oWorkspace.OpenDatabase(txtSecurityDatabaseLocation.Text, True, False, "MS Access;PWD=" & gsSecurityDatabasePassword)
    If Err.Number > 0 Then
        MsgBox "The database could not be opened.", vbInformation, "MACRO"
    Else
        If chkProtectSecurityDatabase.Value = vbChecked Then
            ' NCJ - We are setting the password so assume it's currently ""
'            oDatabase.NewPassword gsSecurityDatabasePassword, gsSecurityDatabasePassword
'            If Err.Number = 0 Then
'                blnChanged = True
'            End If
            oDatabase.NewPassword "", gsSecurityDatabasePassword
            If Err.Number = 0 Then
                blnChanged = True
            Else
                ' Try with password
                oDatabase.NewPassword gsSecurityDatabasePassword, gsSecurityDatabasePassword
                If Err.Number = 0 Then
                    blnChanged = True
                End If
            End If
        ElseIf chkProtectSecurityDatabase.Value = vbUnchecked Then
            ' NCJ - We are unsetting the password so assume it's currently set
            oDatabase.NewPassword gsSecurityDatabasePassword, ""
            If Err.Number = 0 Then
                blnChanged = True
            Else
                ' Try without password
                oDatabase.NewPassword "", ""
                If Err.Number = 0 Then
                    blnChanged = True
                End If
            End If
'            oDatabase.NewPassword "", ""
'            If Err.Number = 0 Then
'                blnChanged = True
'            End If
        End If
        If blnChanged = False Then
            MsgBox "The database password could not be changed.", vbInformation, "MACRO"
        End If
        
        oDatabase.Close
        Set oDatabase = Nothing
    End If
    
'    oDatabase.Close
    Set oDatabase = Nothing
    oWorkspace.Close
    Set oWorkspace = Nothing

    InitializeSecurityADODBConnection

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "chkProtectSecurityDatabase_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'----------------------------------------------------------------------------------------'
Private Function IsSecurityDBProtected(sSDBPath As String) As Boolean
'----------------------------------------------------------------------------------------'
' See if given Security DB is protected
' Try opening with no password and see if it succeeds
' Assume it's a valid MACRO Security DB
'----------------------------------------------------------------------------------------'
Dim oWorkspace As dao.Workspace
Dim oDatabase As dao.Database

    ' Temporarily shut down our own security DB connection
    TerminateSecurityADODBConnection
    
    '   Open the database in exclusive mode
    Set oWorkspace = DBEngine.Workspaces(0)
    
    On Error GoTo ErrProtected
    ' Try opening with NO password and see if there is an error
    Set oDatabase = oWorkspace.OpenDatabase(sSDBPath, True, False, _
                        "MS Access;PWD=" & "")

    ' Seemed OK so assume no password is in place
    IsSecurityDBProtected = False
    
    oDatabase.Close
    oWorkspace.Close
    Set oDatabase = Nothing
    Set oWorkspace = Nothing
    
    ' Reinitialise our own security connection
    InitializeSecurityADODBConnection
    
    Exit Function
    
ErrProtected:
    ' Assume the error is because of the password
    ' (but it might not be...)
    oWorkspace.Close
    Set oDatabase = Nothing
    Set oWorkspace = Nothing
    
    IsSecurityDBProtected = True
    
    ' Reinitialise our own security connection
    InitializeSecurityADODBConnection

End Function

'----------------------------------------------------------------------------------------'
Private Sub cmdBrowser_Click()
'----------------------------------------------------------------------------------------'
' Let them browse for a new security database
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler

    dlgOpenFile.flags = 0
    dlgOpenFile.CancelError = True
    dlgOpenFile.InitDir = txtSecurityDatabaseLocation.Text  'SDM 15/02/00 SR2968
    dlgOpenFile.DefaultExt = "mdb"
    dlgOpenFile.Filter = ".mdb"
    dlgOpenFile.ShowOpen
    dlgOpenFile.flags = cdlOFNFileMustExist
    
    If Err.Number <> cdlCancel Then
        ' Let the txtSecurityDatabaseLocation_Change event
        ' handle the setting of the new security path
       txtSecurityDatabaseLocation.Text = dlgOpenFile.FileName
       'ATO 10/06/2002
       'assigns security database for next macro session to
       'form level variable for display (bug 2.2.8 no.8)
       msNextSecurityDatabase = dlgOpenFile.FileName
    End If
    
    Exit Sub
ErrHandler:
         Select Case Err.Number
            Case cdlCancel
                    Exit Sub
        End Select

End Sub

'------------------------------------------------------------------------------------'
Private Sub cmdCancel_Click()
'------------------------------------------------------------------------------------'
' cancel the form
'------------------------------------------------------------------------------------'
    
    Unload Me

End Sub

Private Sub cmdCreateDB_Click()
    ' Go to New Database form
    frmNewDatabase.Show vbModal
End Sub


'----------------------------------------------------------------------------------------'
Private Sub cmdDefaultSecurity_Click()
'----------------------------------------------------------------------------------------'
' Revert back to the MACRO "default" security database
' Check that it's still a valid security database
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler

'    If txtSecurityDatabaseLocation.Text = "" Then
'        MsgBox "Please browse to a database.", vbInformation, "MACRO"
'        Exit Sub
'    End If

    If IsDatabaseSecurityDB(DefaultSecurityDatabasePath) = False Then
        MsgBox "The default MACRO security database is no longer valid", vbOKOnly, "MACRO"
    Else
        ' Let txtSecurityDatabaseLocation_Change event set the new path
        txtSecurityDatabaseLocation.Text = DefaultSecurityDatabasePath
        'ATO 10/06/2002
        msNextSecurityDatabase = DefaultSecurityDatabasePath
    End If
    
    
    Exit Sub
ErrHandler:
        Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdDefaultSecurity")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

'------------------------------------------------------------------------------------'
Private Sub cmdDBPassword_Click()
'------------------------------------------------------------------------------------'
' choose the correct form to display depending on which option button has been clicked
' strip the last 4 characters out for the database description using replace and right
'------------------------------------------------------------------------------------'
Dim sDescription As String
Dim sPath As String
Dim oWorkspace As Workspace
Dim oDatabase As Database
Dim sSQL As String
Dim rsPassword As ADODB.Recordset
Dim sPassword As String
Dim sMessage As String

    On Error GoTo ErrHandler
  
         frmChangeDatabasePassword.OldPassword = ""
         frmChangeDatabasePassword.NewPassword = ""
        ' Me.Hide
         frmChangeDatabasePassword.Show vbModal
         ' NCJ SR1225 - Check for no change or Cancel
         ' NB If user clicked Cancel old and new passwords will be empty
         If frmChangeDatabasePassword.NewPassword = frmChangeDatabasePassword.OldPassword Then
            Load frmDatabases
            ' Me.Show
             Exit Sub
         End If
    
            If lvwDatabases.SelectedItem.SubItems(1) = "Access" Then
                sDescription = lvwDatabases.SelectedItem
                sPath = Mid(lvwDatabases.SelectedItem.SubItems(2), 10)
                'WillC 28/3/00  changed the closing of the connection
                TerminateMacroADODBConnection
                '   Open the database in exclusive mode
                Set oWorkspace = DBEngine.Workspaces(0)
                
                On Error Resume Next
                Set oDatabase = oWorkspace.OpenDatabase(sPath, True, False, "MS Access;PWD=" & frmChangeDatabasePassword.OldPassword)
                If Err.Number > 0 Then
                    MsgBox "The database could not be opened.", vbInformation, "MACRO"
                Else
                    oDatabase.NewPassword frmChangeDatabasePassword.OldPassword, frmChangeDatabasePassword.NewPassword
                    If Err.Number > 0 Then
                        MsgBox "The database password could not be changed.", vbInformation, "MACRO"
                    Else
                        'Mo Morris 20/9/01 Db Audit (DatabaseDescription to DatabaseCode)
                        sSQL = "UPDATE Databases SET DatabasePassword = '" & frmChangeDatabasePassword.NewPassword & "'" _
                            & " WHERE DatabaseCode = '" & sDescription & "'"
                        SecurityADODBConnection.Execute sSQL
                        MsgBox "The database password has been successfully changed.", vbInformation, "MACRO"
                    End If
                    oDatabase.Close
                    Set oDatabase = Nothing
             End If
             oWorkspace.Close
             Set oWorkspace = Nothing

            On Error GoTo ErrHandler

            If sDescription = goUser.Database.NameOfDatabase Then
                goUser.SetCurrentDatabase goUser.UserName, sDescription, "", "", True, True, sMessage
            End If
            InitializeMacroADODBConnection
        Else
            'MsgBox "This is not an Access database", vbOKOnly + vbInformation, "MACRO"
            ' WillC 8/5/00 SR3336 Allow users to change the password for now Access database.
            'Mo Morris 20/9/01 Db Audit (DatabaseDescription to DatabaseCode)
            sSQL = " Select DatabasePassword from Databases where DatabaseCode = '" & lvwDatabases.SelectedItem & "'"
            Set rsPassword = New ADODB.Recordset
            rsPassword.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
            sPassword = rsPassword!DatabasePassword
            Set rsPassword = Nothing
            
            If LCase(frmChangeDatabasePassword.OldPassword) <> LCase(sPassword) Then
                MsgBox "The database password is not correct. Please try again.", vbInformation, "MACRO"
                frmChangeDatabasePassword.OldPassword = vbNullString
                frmChangeDatabasePassword.Show vbModal
            Else
                'Mo Morris 20/9/01 Db Audit (DatabaseDescription to DatabaseCode)
                sSQL = "UPDATE Databases SET DatabasePassword = '" & frmChangeDatabasePassword.NewPassword & "'" _
                     & " WHERE DatabaseCode = '" & lvwDatabases.SelectedItem & "'"
                SecurityADODBConnection.Execute sSQL
                MsgBox "The database password has been successfully changed.", vbInformation, "MACRO"
            End If
        End If
    'End If
    RefreshDatabases
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdDBPassword_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'------------------------------------------------------------------------------------'
Private Sub cmdOK_Click()
'------------------------------------------------------------------------------------'

'------------------------------------------------------------------------------------'

     Unload Me
    
End Sub

Private Sub cmdRegisterDB_Click()
    'Go to Register Database form
    frmRegisterDatabase.Show vbModal
End Sub

Private Sub Command1_Click()

End Sub

'------------------------------------------------------------------------------------'
Private Sub Form_Load()
'------------------------------------------------------------------------------------'
' Check  to see that the user has the correct rights
'------------------------------------------------------------------------------------'
    
Dim clmX As MSComctlLib.ColumnHeader
    
    Me.Icon = frmMenu.Icon
    
    mbProgramChecking = False
    
    CheckUserDatabaseRights
    
   
    ' Add ColumnHeaders with appropriate widths to lvwDatabases
    Set clmX = lvwDatabases.ColumnHeaders.Add(, , "Database", (lvwDatabases.Width - 110) * 0.2)
    Set clmX = lvwDatabases.ColumnHeaders.Add(, , "Type", (lvwDatabases.Width - 110) * 0.2)
    Set clmX = lvwDatabases.ColumnHeaders.Add(, , "Parameters", (lvwDatabases.Width - 110) * 3)
    
    RefreshDatabases
  
    
End Sub

'------------------------------------------------------------------------------------'
Public Sub CheckUserDatabaseRights()
'------------------------------------------------------------------------------------'
' Check  to see that the user has the correct rights to create or register db's
'------------------------------------------------------------------------------------'
    
    If goUser.CheckPermission(gsFnRegisterDB) = True Then
        cmdRegisterDb.Enabled = True
        cmdDBPassword.Enabled = True
        Frame2.Enabled = True
    Else
        cmdRegisterDb.Enabled = False
        cmdDBPassword.Enabled = False
        Frame2.Enabled = False
    End If
    
    
    If goUser.CheckPermission(gsFnCreateDB) = True Then
        cmdCreateDB.Enabled = True
    Else
        cmdCreateDB.Enabled = False
    End If
    
End Sub

'----------------------------------------------------------------------------------------------
Public Sub RefreshDatabases()
'----------------------------------------------------------------------------------------------
' Add the databases to their listview
'----------------------------------------------------------------------------------------------

 'Create a variable to add ListItem objects and receive the list of databases.
Dim itmX As MSComctlLib.ListItem
Dim sSQL As String
Dim rsDatabaseList As ADODB.Recordset

    On Error GoTo ErrHandler

    sSQL = "SELECT * FROM Databases"
    Set rsDatabaseList = New ADODB.Recordset
    rsDatabaseList.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, , adCmdText

    lvwDatabases.ListItems.Clear
    
    ' While the record is not the last record, add a ListItem object.
    With rsDatabaseList
        Do Until .EOF = True
            'this places the databasenames in the listview
            'Mo Morris 20/9/01 Db Audit (DatabaseDescription to DatabaseCode)
            Set itmX = lvwDatabases.ListItems.Add(, , !DataBaseCode)
            Select Case !DatabaseType
            Case MACRODatabaseType.Access
                itmX.SubItems(1) = "Access"
                itmX.SubItems(2) = "Location=" & !DatabaseLocation
            Case MACRODatabaseType.sqlserver
                itmX.SubItems(1) = "SQL Server 6.5"
                itmX.SubItems(2) = "Server=" & !ServerName & ";Database=" & !NameOfDatabase
            Case MACRODatabaseType.SQLServer70
                itmX.SubItems(1) = "SQL Server 7.0"
                itmX.SubItems(2) = "Server=" & !ServerName & ";Database=" & !NameOfDatabase
            Case MACRODatabaseType.Oracle80
                itmX.SubItems(1) = "Oracle 8.0"
                itmX.SubItems(2) = "Server=" & !ServerName & ";Database=" & !NameOfDatabase
            End Select
            .MoveNext   ' Move to next record.
            
        Loop
    End With
   
    rsDatabaseList.Close
    Set rsDatabaseList = Nothing
   
    'ATO 10/06/2002
    txtCurrentSecurityDatabaseLocation.Text = SecurityDatabasePath
   
    If msNextSecurityDatabase = "" Then
        txtSecurityDatabaseLocation.Text = SecurityDatabasePath
    Else
        txtSecurityDatabaseLocation.Text = msNextSecurityDatabase
    End If
    
    txtSecurityDatabaseLocation.Enabled = False

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "RefreshDatabases")
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
Private Sub txtSecurityDatabaseLocation_Change()
'----------------------------------------------------------------------------------------------
' Check for valid Security database
' then store in Registry
' NCJ 24/1/00 - Use new SecurityDatabasePath property of gUser
'----------------------------------------------------------------------------------------------
Dim sNewSecurityPath As String

    On Error GoTo ErrHandler
    
    sNewSecurityPath = txtSecurityDatabaseLocation.Text
    
    If sNewSecurityPath > "" Then
        If IsDatabaseSecurityDB(sNewSecurityPath) Then
            SecurityDatabasePath = sNewSecurityPath
            ' NCJ - Say that it's a programmatic check box click
            mbProgramChecking = True
            If IsSecurityDBProtected(sNewSecurityPath) Then
                chkProtectSecurityDatabase.Value = 1
            Else
                chkProtectSecurityDatabase.Value = 0
            End If
            ' Reset value in case we're in form load
            ' and the check box doesn't change
            mbProgramChecking = False
        Else
            ' Reset to empty string
            txtSecurityDatabaseLocation.Text = SecurityDatabasePath
'            txtSecurityDatabaseLocation.Text = ""
            ' Value of 2 makes it greyed out
            If IsSecurityDBProtected(SecurityDatabasePath) Then
                chkProtectSecurityDatabase.Value = 1
            Else
                chkProtectSecurityDatabase.Value = 0
            End If
            'chkProtectSecurityDatabase.Value = 2
        End If
    End If
Exit Sub

ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtSecurityDatabaseLocation_Change")
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
Private Function IsDatabaseSecurityDB(sDatabasePath As String) As Boolean
'----------------------------------------------------------------------------------------------
' Before you can set a database as the default database first check to make sure its a
' Security database.
' NCJ 24 Jan 00 - Added argument sDataBasePath
'----------------------------------------------------------------------------------------------

Dim ADODBConnection As ADODB.Connection
Dim rsSecurityTest As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrNotMacro
    
    ' Temporarily shut down our own security DB connection
    TerminateSecurityADODBConnection
    
    IsDatabaseSecurityDB = False
    
    Set ADODBConnection = New ADODB.Connection

'   ATN 24/2/2000 SR 3096
'   Security database is always Access (and gUser.DatabaseType doesn't refer to the
'   security database anyway) so always open using Access provider
'    Select Case gUser.DatabaseType
'    Case MACRODatabaseType.Access
            'JET OLE DB NATIVE PROVIDER
            ADODBConnection.Open Connection_String(CONNECTION_MSJET_OLEDB_40, sDatabasePath, , , gsSecurityDatabasePassword)
'    Case MACRODatabaseType.SQLServer
'            'SQL SERVER OLE DB NATIVE PROVIDER
'            ADODBConnection.Open "PROVIDER=SQLOLEDB;" & _
'                "DATA SOURCE=" & gUser.ServerName & ";DATABASE=" & gUser.DatabaseName & ";" & _
'                "USER ID=" & gUser.DatabaseUser & ";PASSWORD=" & gUser.DatabasePassword
'   ATN 21/12/99
'   Added connection to Oracle
'    Case MACRODatabaseType.Oracle80
'            'Oracle OLE DB NATIVE PROVIDER
'            ADODBConnection.Open "PROVIDER=MSDAORA;" & _
'                "DATA SOURCE=;" & _
'                "USER ID=" & gUser.DatabaseUser & ";PASSWORD=" & gUser.DatabasePassword
'    End Select
    
    Set rsSecurityTest = New ADODB.Recordset
    ' Try and read the "MacroUser" table
    sSQL = "Select * from MacroUser"
    rsSecurityTest.Open sSQL, ADODBConnection, adOpenKeyset, , adCmdText
    
    If rsSecurityTest.RecordCount > 0 Then
        IsDatabaseSecurityDB = True
    End If
    
    On Error GoTo ErrHandler
    
    rsSecurityTest.Close
    ADODBConnection.Close
    Set rsSecurityTest = Nothing
    Set ADODBConnection = Nothing
    
    InitializeSecurityADODBConnection
    
 Exit Function
 
ErrNotMacro:
    ' NCJ - We don't know exactly what errors we might get
    ' (or what their error numbers are...)
'      If Err.Number = -2147217865 Then
            MsgBox "The database you have chosen has not been recognised " & vbCrLf _
                 & "as a valid MACRO Security database.", vbInformation, "MACRO"

'            rsSecurityTest.Close
'            ADODBConnection.Close
            Set rsSecurityTest = Nothing
            Set ADODBConnection = Nothing
             
            ' Restart our own security DB connection
            InitializeSecurityADODBConnection

'     End If
     Exit Function
     
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "IsDatabaseSecurityDB")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Function
