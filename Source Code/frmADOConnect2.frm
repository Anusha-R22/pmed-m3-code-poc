VERSION 5.00
Begin VB.Form frmADOConnect2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connection String"
   ClientHeight    =   2295
   ClientLeft      =   6570
   ClientTop       =   4470
   ClientWidth     =   4710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdTest 
      Caption         =   "&Test"
      Height          =   375
      Left            =   60
      TabIndex        =   5
      Top             =   1860
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3420
      TabIndex        =   7
      Top             =   1860
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2100
      TabIndex        =   6
      Top             =   1860
      Width           =   1215
   End
   Begin VB.Frame fraStep3 
      Caption         =   "Connection Values"
      Height          =   1800
      Left            =   60
      TabIndex        =   11
      Top             =   0
      Width           =   4600
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
         Top             =   690
         Width           =   2800
      End
      Begin VB.TextBox txtDatabase 
         Height          =   300
         Left            =   1700
         TabIndex        =   3
         Top             =   1020
         Width           =   2800
      End
      Begin VB.TextBox txtServer 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1700
         TabIndex        =   4
         Top             =   1350
         Width           =   2800
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
         Left            =   135
         TabIndex        =   8
         Top             =   690
         Width           =   735
      End
      Begin VB.Label lblDatabase 
         AutoSize        =   -1  'True
         Caption         =   "Data&base:"
         Height          =   195
         Left            =   135
         TabIndex        =   9
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label lblServer 
         AutoSize        =   -1  'True
         Caption         =   "&Server:"
         Height          =   195
         Left            =   135
         TabIndex        =   10
         Top             =   1350
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmADOConnect2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   File:       frmADOConnect2.frm
'   Author:     Mo Morris 31 August 2001
'   Purpose:    This is a copy of form frmADOConnect with registration code
'               removed.
'               It is in module MACRO_SD for the purpose of connnecting
'               to an SQL Server database for the retreval of category codes and values.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Revisions:
'
'------------------------------------------------------------------------------
Option Explicit

Private msConnectString As String
Private mnDatabaseType As Integer
Private msFormUsage As String

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
   
    Call GetConnectString

    txtUID.Text = ""
    txtPWD.Text = ""
    txtDatabase.Text = ""
    txtServer.Text = ""
    
    Me.Hide
    
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
    
    Me.Icon = frmMenu.Icon
    
    txtUID.Text = ""
    txtPWD.Text = ""
    txtDatabase.Text = ""
    txtServer.Text = ""
    
    cmdOK.Enabled = False
       
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
        txtServer.Enabled = True
    End Select

    FormCentre Me

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
Private Function CheckForMacroTables(ByVal oConnection As ADODB.Connection, ByRef sVersion As String) As Boolean
'------------------------------------------------------------------------------
'Check there is not already a database
'
'Mo Morris  17/8/01 This function no longer displays a message.
'                   It now updates SVersion as well as return a Boolean (success/fail)
'------------------------------------------------------------------------------
Dim rs As ADODB.Recordset
Dim sMessage As String

    On Error GoTo ErrDBExists
    Set rs = New ADODB.Recordset
    'next line causes an error and jumps to the error handler if there is no MACROControl table
    rs.Open "SELECT MACROVersion, BuildSubVersion FROM MACROCONTROL", oConnection
    sVersion = rs!MACROVersion & "." & rs!BuildSubVersion
    
    rs.Close
    Set rs = Nothing
    
    CheckForMacroTables = True
    Exit Function
    
ErrDBExists:
    Set rs = Nothing
    CheckForMacroTables = False
    Exit Function

End Function

'------------------------------------------------------------------------------
Private Sub TestConnectString()
'------------------------------------------------------------------------------
Dim oCon As ADODB.Connection
Dim sConnectString As String
Dim bDatabaseExists As Boolean
Dim sVersion As String
Dim sErr As String
Dim lErrNo As Long

    On Error GoTo ErrHandler

    Screen.MousePointer = vbHourglass
  
    'Create a connection to the soon to be created database
    Set oCon = New ADODB.Connection
    Select Case mnDatabaseType
    Case MACRODatabaseType.Oracle80
         sConnectString = Connection_String(CONNECTION_MSDAORA, txtDatabase.Text, , txtUID.Text, txtPWD.Text)
    Case MACRODatabaseType.sqlserver, MACRODatabaseType.SQLServer70
        sConnectString = Connection_String(CONNECTION_SQLOLEDB, txtServer.Text, txtDatabase.Text, txtUID.Text, txtPWD.Text)
    End Select
    oCon.ConnectionString = sConnectString
    
    On Error Resume Next
    oCon.Open sConnectString
    sErr = Err.Description
    lErrNo = Err.Number
    Err.Clear
    On Error GoTo ErrHandler
  
    'Check the connection state
    'ASH 10/12/2002 Display more appropriate messages should errors occur during registration
    If lErrNo <> 0 Then
        Screen.MousePointer = vbDefault
        DialogInformation ("The connection to the specified database failed because of the following:") & vbCrLf _
        & sErr
        cmdOK.Enabled = False
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        Call DialogInformation("Connection successful.", "Connect")
    End If

'    'Check if the database already exists
'    sVersion = ""
'    bDatabaseExists = CheckForMacroTables(sConnect, sVersion)
'
'
'    If bDatabaseExists Then
'        'Database is MACRO
'        DialogInformation ("The connection details you have provided refer to a MACRO database (Version " & sVersion & ").")
'
'    Else
'        'Database is not MACRO
'        DialogError ("The connection details you have provided do not refer to a database that contains MACRO tables.")
'    End If

    'Enable OK button
    cmdOK.Enabled = True
    
    oCon.Close
    Set oCon = Nothing
    
    Screen.MousePointer = vbDefault
 
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

'------------------------------------------------------------------------------
Private Sub txtDatabase_Change()
'------------------------------------------------------------------------------

    cmdOK.Enabled = False

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

