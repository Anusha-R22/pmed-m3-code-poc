VERSION 5.00
Begin VB.Form frmADOConnect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create or Register"
   ClientHeight    =   2790
   ClientLeft      =   6570
   ClientTop       =   4470
   ClientWidth     =   4710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdTest 
      Caption         =   "&Test"
      Height          =   375
      Left            =   60
      TabIndex        =   6
      Top             =   2340
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3420
      TabIndex        =   8
      Top             =   2340
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2100
      TabIndex        =   7
      Top             =   2340
      Width           =   1215
   End
   Begin VB.Frame fraStep3 
      Caption         =   "Connection Values"
      Height          =   2280
      Left            =   60
      TabIndex        =   12
      Top             =   0
      Width           =   4600
      Begin VB.TextBox txtAlias 
         Height          =   330
         Left            =   1700
         MaxLength       =   15
         TabIndex        =   5
         Top             =   1800
         Width           =   2800
      End
      Begin VB.TextBox txtUID 
         Height          =   300
         Left            =   1700
         TabIndex        =   1
         Top             =   360
         Width           =   2800
      End
      Begin VB.TextBox txtPWD 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1700
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   720
         Width           =   2800
      End
      Begin VB.TextBox txtDatabase 
         Height          =   300
         Left            =   1700
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1065
         Width           =   2800
      End
      Begin VB.TextBox txtServer 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1700
         TabIndex        =   4
         Top             =   1425
         Width           =   2800
      End
      Begin VB.Label lblAlias 
         Caption         =   "&MACRO DB Alias:"
         Height          =   255
         Left            =   135
         TabIndex        =   13
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblUID 
         AutoSize        =   -1  'True
         Caption         =   "UID:"
         Height          =   195
         Left            =   135
         TabIndex        =   0
         Top             =   360
         Width           =   330
      End
      Begin VB.Label lblPWD 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblDatabase 
         AutoSize        =   -1  'True
         Caption         =   "Data&base:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblServer 
         AutoSize        =   -1  'True
         Caption         =   "&Server:"
         Height          =   195
         Left            =   135
         TabIndex        =   11
         Top             =   1440
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmADOConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2000-2004. All Rights Reserved
'   File:       frmADOConnect.frm
'   Author:     Will Casey, Jan 5 2000
'   Purpose:    To allow the user to connect to Ado Datasources (Oracle and
'               SQL Server) for the purpose of Creating and Registering databases.
'
'               The registration task used to be handled by frmSQLDBRegister
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Revisions:
'Mo 13/8/01 Re-write stemming from the need to rationalise/standardise around
'           db connection strings as well as paramaterizing the calls to this form
'           for the separate tasks of DB creation and DB registration.
'           Unneccessary controls (text boxes and labels) removed.
'           Previous revision history removed.
'           Old/Unused code removed
'ASH 15/04/2002 - Made changes to routine RegisterThisDatabase
'ASH 12/9/2002 - Minor change to RegisterThisDb
'ASH 10/10/2002 changed from txtDatabase to txtAlias in routine RegisterThisDb
'REM 23/10/02 - check to see if user has entered an Alias, if not, use database name in routine RegisterThisDb
'ASH 4/12/2002 check for invalid characters in TestConnectString
'ASH 6/12/2002 Defaulted HTML Location
'REM 10/02/04 - Added routine EnableTestButton
'MLM 20/06/05: bug 2556: allow db names > 15 chars
'------------------------------------------------------------------------------
Option Explicit

Private msConnectString As String
Private mnDatabaseType As Integer
Private msFormUsage As String
Private mbIsLoading As Boolean
Private mbUserChangedAlias As Boolean

Public Property Get ConnectString() As String

    ConnectString = msConnectString

End Property

'------------------------------------------------------------------------------
Private Sub cmdCancel_Click()
'------------------------------------------------------------------------------
' if cancel is clicked then set the connect string to nothing
'------------------------------------------------------------------------------
   
    msConnectString = vbNullString
   
    txtUID.Text = ""
    txtPWD.Text = ""
    txtDatabase.Text = ""
    txtServer.Text = ""
    
    Me.Hide

End Sub

'------------------------------------------------------------------------------
Private Sub cmdOK_Click()
'------------------------------------------------------------------------------
' get the connection string and pass it back to frmNewdatabase
'------------------------------------------------------------------------------
On Error GoTo ErrHandler
   
    If cmdOK.Caption = "Create" Then
        Call GetConnectString
    Else
        Call RegisterThisDb
    End If
        
    txtUID.Text = ""
    txtPWD.Text = ""
    txtDatabase.Text = ""
    txtServer.Text = ""
    txtAlias.Text = ""
    
    Me.Hide
    
    DoEvents
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdOK_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'------------------------------------------------------------------------------
Private Sub cmdTest_Click()
'------------------------------------------------------------------------------
' Allow the user to test the connection string
'------------------------------------------------------------------------------

    Call TestConnectString
      
End Sub

'------------------------------------------------------------------------------
Private Sub Form_Load()
'------------------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    mbIsLoading = True
    
    Me.Icon = frmMenu.Icon
    
    mbUserChangedAlias = False
    
    txtUID.Text = ""
    txtPWD.Text = ""
    txtDatabase.Text = ""
    txtServer.Text = ""
    
    cmdOK.Enabled = False
    
    Select Case msFormUsage
    Case "Create"
        cmdOK.Caption = "Create"
        txtAlias.Enabled = False
        Select Case mnDatabaseType
            Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
                Me.Caption = msFormUsage & " SQL Server Database"
                lblDatabase.Caption = "Database Name:"
                lblServer.Visible = True
                txtServer.Visible = True
                txtServer.Enabled = True
                lblAlias.Visible = False
                txtAlias.Visible = False
                fraStep3.Height = fraStep3.Height - txtAlias.Height
                cmdTest.Top = fraStep3.Height + 100
                cmdCancel.Top = fraStep3.Height + 100
                cmdOK.Top = fraStep3.Height + 100
                frmADOConnect.Height = frmADOConnect.Height - 300
            Case MACRODatabaseType.Oracle80
                Me.Caption = msFormUsage & " Oracle Database"
                lblDatabase.Caption = "Net Service Name:"
                lblServer.Visible = False
                txtServer.Visible = False
                lblAlias.Visible = False
                txtAlias.Visible = False
                fraStep3.Height = fraStep3.Height - (txtAlias.Height + txtServer.Height)
                cmdTest.Top = fraStep3.Height + 100
                cmdCancel.Top = fraStep3.Height + 100
                cmdOK.Top = fraStep3.Height + 100
                frmADOConnect.Height = frmADOConnect.Height - 600
        End Select
    Case "Register"
        cmdOK.Caption = "Register"
        Select Case mnDatabaseType
            Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
                Me.Caption = msFormUsage & " SQL Server Database"
                lblDatabase.Caption = "Database Name:"
                lblServer.Visible = True
                txtServer.Visible = True
                txtServer.Enabled = True
            Case MACRODatabaseType.Oracle80
                Me.Caption = msFormUsage & " Oracle Database"
                lblDatabase.Caption = "Net Service Name:"
                lblServer.Visible = False
                txtServer.Visible = False
                txtAlias.Top = txtServer.Top
                lblAlias.Top = lblServer.Top
                fraStep3.Height = fraStep3.Height - txtServer.Height
                cmdTest.Top = fraStep3.Height + 100
                cmdCancel.Top = fraStep3.Height + 100
                cmdOK.Top = fraStep3.Height + 100
                frmADOConnect.Height = frmADOConnect.Height - 300

        End Select
    End Select

    Call EnableTestButton(mnDatabaseType)

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

'-------------------------------------------------------------------------
Private Sub EnableTestButton(nDatabaseType As Integer)
'-------------------------------------------------------------------------
'REM 10/02/04 - Don't enable Test button fro SQL Server unless there is something
' in the server and database fields
'-------------------------------------------------------------------------
    
    If nDatabaseType = MACRODatabaseType.Oracle80 Then
        'Don't have to worry about disableing the Test button as it will take care of itself
        cmdTest.Enabled = True
    ElseIf nDatabaseType = MACRODatabaseType.sqlserver Then
        If (Trim(txtServer.Text) <> "") And (Trim(txtDatabase.Text) <> "") Then
            cmdTest.Enabled = True
        Else
            cmdTest.Enabled = False
        End If
        
    End If
     
End Sub

'------------------------------------------------------------------------------
Public Sub GetConnectString()
'------------------------------------------------------------------------------
' build the connection string and place it in the msConnectString variable
'------------------------------------------------------------------------------
Dim sConnect As String

    On Error GoTo ErrHandler
    
    Select Case Me.DatabaseType
    Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
        msConnectString = Connection_String(CONNECTION_SQLOLEDB, txtServer.Text, txtDatabase.Text, txtUID.Text, txtPWD.Text)
    Case MACRODatabaseType.Oracle80
        msConnectString = Connection_String(CONNECTION_MSDAORA, txtDatabase.Text, , txtUID.Text, txtPWD.Text)
    End Select
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "GetConnectString")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'------------------------------------------------------------------------------
Private Sub TestConnectString()
'------------------------------------------------------------------------------
Dim oCon As ADODB.Connection
Dim sConnectString As String
Dim bDatabaseExists As Boolean
Dim sVersion As String
Dim lErrNo As Long
Dim sErr As String

    On Error GoTo ErrHandler

    HourglassOn
    
    'ASH 4/12/2002 check for invalid characters
    If Not gblnValidString(txtAlias.Text, valAlpha + valNumeric + valSpace + valUnderscore + valDecimalPoint) Then
        Screen.MousePointer = vbDefault
        Call DialogInformation("Database alias contains invalid characters")
        Exit Sub
    ElseIf Not gblnValidString(txtUID.Text, valAlpha + valNumeric + valSpace + valUnderscore + valDecimalPoint) Then
        Screen.MousePointer = vbDefault
        Call DialogInformation("User ID contains invalid characters")
        Exit Sub
    ElseIf Not gblnValidString(txtDatabase.Text, valAlpha + valNumeric + valSpace + valUnderscore + valDecimalPoint) Then
        Screen.MousePointer = vbDefault
        Call DialogInformation("Net service name contains invalid characters")
        Exit Sub
    End If

    'Create a connection to the soon to be created database
    Set oCon = New ADODB.Connection
    Select Case mnDatabaseType
    Case MACRODatabaseType.Oracle80
         sConnectString = Connection_String(CONNECTION_MSDAORA, Trim(txtDatabase.Text), , Trim(txtUID.Text), Trim(txtPWD.Text))
    Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
        'REM 01/04/03 - Create connection without database name
        sConnectString = Connection_String(CONNECTION_SQLOLEDB, Trim(txtServer.Text), "", Trim(txtUID.Text), Trim(txtPWD.Text))
        
        On Error Resume Next
        oCon.Open sConnectString
        oCon.Execute "Create database " & Trim(txtDatabase.Text)
        oCon.Close
        
        On Error GoTo ErrHandler
        sConnectString = Connection_String(CONNECTION_SQLOLEDB, Trim(txtServer.Text), Trim(txtDatabase.Text), Trim(txtUID.Text), Trim(txtPWD.Text))
        
    End Select
   
    On Error Resume Next
    oCon.Open sConnectString
    sErr = Err.Description
    lErrNo = Err.Number
    Err.Clear
    
    On Error GoTo ErrHandler
    
    HourglassOff
  
    'Check the connection state
    'ASH 10/12/2002 Display more appropriate messages should errors occur during registration
    If lErrNo <> 0 Then
        Screen.MousePointer = vbDefault
        DialogInformation ("The connection to the specified database failed because of the following:") & vbCrLf _
        & sErr
        cmdOK.Enabled = False
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    On Error GoTo ErrHandler

    'Check if the database already exists
    sVersion = ""
    bDatabaseExists = CheckForMacroTables(oCon, sVersion)
    
    If msFormUsage = "Create" Then
        If bDatabaseExists Then
            'Database exists, (i.e. contains Macro tables), Do not Create
            DialogError ("The connection details you have provided refer to an already existing MACRO database (Version " & sVersion & ")." _
                & vbCr & "Database Creation cannot proceed.")
        Else
            'Database does not contain Macro Tables, OK to Create
            Select Case mnDatabaseType
            Case MACRODatabaseType.Oracle80
                DialogInformation ("The connection to Net Service Name " & txtDatabase.Text & vbCr _
                   & "for user " & txtUID.Text & " using the specified password " & vbCr _
                   & "has tested successfully." & vbCr _
                   & "Proceed with creating the Macro database tables.")
            Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
                DialogInformation ("The connection to " & txtServer.Text & vbCr _
                   & "for database " & txtDatabase.Text & vbCr _
                   & "for user " & txtUID.Text & " using the specified password " & vbCr _
                   & "has tested successfully." & vbCr _
                   & "Proceed with creating the Macro database tables.")
            End Select
            'Enable the 'Create' command button
            cmdOK.Enabled = True
        End If
    Else    'msFormUsage = "Register"
        If bDatabaseExists Then
            'Database exists, OK to Register
            DialogInformation ("The connection details you have provided refer to a MACRO database (Version " & sVersion & ")." _
                & vbCr & "Proceed with Registration.")
            'Enable the 'Register' command button
            cmdOK.Enabled = True
        Else
            'Database does not contain Macro Tables, Do not Register
            DialogError ("The connection details you have provided do not refer to a database that contains MACRO tables." _
                & vbCr & "Registration cannot proceed.")
        End If
    End If
    
    oCon.Close
    Set oCon = Nothing
 
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "TestConnectString")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'------------------------------------------------------------------------------
Public Property Get DatabaseType() As Variant
'------------------------------------------------------------------------------

    DatabaseType = mnDatabaseType

End Property

'------------------------------------------------------------------------------
Public Property Let DatabaseType(ByVal vNewValue As Variant)
'------------------------------------------------------------------------------

    mnDatabaseType = vNewValue

End Property

'----------------------------------------------------------------------------
Private Sub txtAlias_Change()
'----------------------------------------------------------------------------
'
'----------------------------------------------------------------------------


End Sub

'------------------------------------------------------------------------------
Private Sub txtAlias_KeyPress(KeyAscii As Integer)
'------------------------------------------------------------------------------

    mbUserChangedAlias = True
    
End Sub

'------------------------------------------------------------------------------
Private Sub txtDatabase_Change()
'------------------------------------------------------------------------------

    cmdOK.Enabled = False
    
    If Not mbIsLoading Then
        If txtAlias.Text = vbNullString Or Not mbUserChangedAlias Then
            txtAlias.Text = txtDatabase.Text
        End If
        
        Call EnableTestButton(mnDatabaseType)
    End If

End Sub

'------------------------------------------------------------------------------
Private Sub txtPWD_Change()
'------------------------------------------------------------------------------

    cmdOK.Enabled = False

End Sub

'------------------------------------------------------------------------------
Private Sub txtServer_Change()
'------------------------------------------------------------------------------

    cmdOK.Enabled = False
    
    Call EnableTestButton(mnDatabaseType)

End Sub

'------------------------------------------------------------------------------
Private Sub txtUID_Change()
'------------------------------------------------------------------------------

    cmdOK.Enabled = False

End Sub

'------------------------------------------------------------------------------
Public Property Get FormUsage() As String
'------------------------------------------------------------------------------

    FormUsage = msFormUsage

End Property

'------------------------------------------------------------------------------
Public Property Let FormUsage(ByVal sNewValue As String)
'------------------------------------------------------------------------------

    msFormUsage = sNewValue

End Property

'------------------------------------------------------------------------------
Private Sub RegisterThisDb()
'------------------------------------------------------------------------------
' MLM 20/06/05: bug 2556: allow db names > 15 chars
'------------------------------------------------------------------------------
'Dim sSQL As String
'Dim sHTMLPath As String
'Dim sSHTMLPATH As String
Dim sDatabaseAlias As String
Dim sNameofDatabase As String
Dim sDatabasePswd As String
Dim sDatabaseUser As String
'Dim sEncryptedDatabasePswd As String
'Dim sEncryptedDatabaseUser As String
    
    On Error GoTo ErrHandler
       
    'Disallow the use of numbers only as Database names
    If IsNumeric(txtDatabase.Text) Then
        Select Case Me.DatabaseType
            Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
                DialogError ("You may not register a number as a database name.")
                txtDatabase.Text = vbNullString
                txtDatabase.SetFocus
            Case MACRODatabaseType.Oracle80
                DialogError ("You may not register a number as a Net Service name.")
                txtServer.Text = vbNullString
                txtServer.SetFocus
        End Select
        Exit Sub
    End If
     
    ' MLM 20/06/05: bug 2556
'    'Check that the databasename is not greater than 15 characters
'    If Len(txtDatabase.Text) > 15 Then
'        DialogError ("You cannot register a database with a name that is greater than 15 characters long.")
'        Exit Sub
'    End If

    'Register the database by adding it to the Databases table in the Security database
    'Mo Morris 20/9/01 Db Audit (DatabaseDescription to DatabaseCode)
    'ASH Added new field SecureHTMLLocation to SQL
    'ASH 10/10/2002 changed from txtDatabase to txtAlias
    'REM 23/10/02 - check to see if user has entered an Alias, if not use database name
    sDatabaseAlias = Trim(txtAlias.Text)
    sNameofDatabase = Trim(txtDatabase.Text)
    
    If sDatabaseAlias = "" Then
        sDatabaseAlias = Left(sNameofDatabase, 15)
    End If
    
    If DoesDatabaseExist(sDatabaseAlias) = False Then
        sDatabasePswd = Trim(txtPWD.Text)
        sDatabaseUser = Trim(txtUID.Text)
        
        Call GetConnectString
        Call RegisterMACRODatabase(msConnectString, sDatabaseAlias, Me.DatabaseType, txtServer.Text, sNameofDatabase, sDatabasePswd, sDatabaseUser)
    
'    'ASH 15/04/2002 Added new parameter to frmSetHTMLFolder
'        Call frmSetHTMLFolder.HTMLPath(sHTMLPath, sSHTMLPATH)
'        If sHTMLPath = "" Then Exit Sub
'        sDatabasePswd = txtPWD.Text
'        sDatabaseUser = txtUID.Text
'
'        'encrypt the database password and user
'        If sDatabasePswd = "" Then
'            sEncryptedDatabasePswd = "null"
'        Else
'
'            sEncryptedDatabasePswd = "'" & EncryptString(sDatabasePswd) & "'"
'        End If
'
'        If sDatabaseUser = "" Then
'            sEncryptedDatabaseUser = "null"
'        Else
'            sEncryptedDatabaseUser = "'" & EncryptString(sDatabaseUser) & "'"
'        End If
'
'        'ASH 15/04/2002 to get value for new field for SQL string
'        sSQL = "INSERT INTO Databases (DatabaseCode,DatabaseType,ServerName,NameOfDatabase,DatabaseUser,DatabasePassword,HTMLLocation,SecureHTMLLocation)"
'        'ASH 12/9/2002 Added Alias to get new database name
'        sSQL = sSQL & "VALUES ('" & sDatabaseAlias & "', " & Me.DatabaseType & ",'" & txtServer.Text & "','" & sNameofDatabase & "'," & sEncryptedDatabaseUser & "," & sEncryptedDatabasePswd & ", '" & sHTMLPath & "','" & sSHTMLPATH & "'" & ")"
'        SecurityADODBConnection.Execute sSQL, adCmdText
        
'        'REM 11/02/03 - restore User Database links
'        Call RestoreUserDatabaseLinks
'
'        DialogInformation ("The database has been registered successfully")
''        'Refresh the list of databases on frmDatabases
''        frmDatabases.RefreshDatabases
'        Call frmMenu.RefereshDatabaseInfoForm
    Else
        DialogInformation ("A database with the alias " & sDatabaseAlias & " has already been registered!")
    End If

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "RegisterThisDb")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub


