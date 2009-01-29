VERSION 5.00
Begin VB.Form frmActiveDirectoryNewServer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Active Directory Server"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   Icon            =   "frmActiveDirectoryNewServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2460
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   2460
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2940
      TabIndex        =   7
      Top             =   2460
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2235
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5535
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   2
         Top             =   585
         Width           =   4335
      End
      Begin VB.ComboBox cboLoginType 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   180
         Width           =   4335
      End
      Begin VB.TextBox txtPath 
         Height          =   315
         Left            =   1080
         MaxLength       =   255
         TabIndex        =   3
         Top             =   990
         Width           =   4335
      End
      Begin VB.TextBox txtUserName 
         Height          =   315
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1395
         Width           =   4335
      End
      Begin VB.TextBox txtPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1800
         Width           =   4335
      End
      Begin VB.Label lblName 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   645
         Width           =   795
      End
      Begin VB.Label lblLoginType 
         Caption         =   "Login type"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lblPath 
         Caption         =   "Path"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1050
         Width           =   855
      End
      Begin VB.Label lblUserName 
         Caption         =   "User name"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1395
         Width           =   915
      End
      Begin VB.Label lblPassword 
         Caption         =   "Password"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmActiveDirectoryNewServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
'   File:       frmActiveDirectoryNewServer.frm
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   Author:     I Curtis, December 2005
'   Purpose:    Active Directory new server form
'------------------------------------------------------------------------------
' REVISIONS
'------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------
Private Sub cboLoginType_Click()
'---------------------------------------------------------------------
' ic 07/12/2005
' login type change
'---------------------------------------------------------------------
    Select Case cboLoginType.ListIndex
    Case 0:
        lblPath.Enabled = False
        txtPath.Enabled = False
        txtPath.Text = ""
        lblUserName.Enabled = False
        txtUserName.Enabled = False
        txtUserName.Text = ""
        lblPassword.Enabled = False
        txtPassword.Enabled = False
        txtPassword.Text = ""
    Case 1:
        lblPath.Enabled = True
        txtPath.Enabled = True
        lblUserName.Enabled = False
        txtUserName.Enabled = False
        txtUserName.Text = ""
        lblPassword.Enabled = False
        txtPassword.Enabled = False
        txtPassword.Text = ""
    Case 2:
        lblPath.Enabled = True
        txtPath.Enabled = True
        lblUserName.Enabled = True
        txtUserName.Enabled = True
        lblPassword.Enabled = True
        txtPassword.Enabled = True
    Case 3:
        lblPath.Enabled = True
        txtPath.Enabled = True
        lblUserName.Enabled = False
        txtUserName.Enabled = False
        txtUserName.Text = ""
        lblPassword.Enabled = False
        txtPassword.Enabled = False
        txtPassword.Text = ""
    End Select
End Sub

'---------------------------------------------------------------------
Private Sub cmdCancel_Click()
'---------------------------------------------------------------------
' ic 06/12/2005
' cancel
'---------------------------------------------------------------------
    Unload Me
End Sub

'---------------------------------------------------------------------
Private Sub cmdOK_Click()
'---------------------------------------------------------------------
' ic 06/12/2005
' add new server
'---------------------------------------------------------------------
Dim oCon As ADODB.Connection
Dim sSQL As String
Dim nNextConnectionOrder  As Integer


    On Error GoTo ErrLabel
        
    Select Case cboLoginType.ListIndex
    Case 0:
    Case 1:
        If (txtPath.Text = "") Then
            Call DialogWarning("You must enter an Active Directory server Path")
            Exit Sub
        End If
    Case 2:
        If (txtPath.Text = "") Then
            Call DialogWarning("You must enter an Active Directory server Path")
            Exit Sub
        End If
        If (txtUserName.Text = "") Then
            Call DialogWarning("You must enter a username")
            Exit Sub
        End If
        If (txtPassword.Text = "") Then
            Call DialogWarning("You must enter a password")
            Exit Sub
        End If
    Case 3:
        If (txtPath.Text = "") Then
            Call DialogWarning("You must enter an Active Directory server Path")
            Exit Sub
        End If
    End Select
    
    
    HourglassOn
    
    'get next sequence
    nNextConnectionOrder = NextConnectionOrder
    
    'open db connection
    Set oCon = New ADODB.Connection
    oCon.Open (SecurityADODBConnection)
            
    'insert the new server row
    sSQL = "INSERT INTO ACTIVEDIRECTORYSERVERS (CONNECTORDER, PATH, USERNAME, PASSWORD, LOGINTYPE, NAME) VALUES ("
    sSQL = sSQL & nNextConnectionOrder & ", "
    sSQL = sSQL & IIf((txtPath.Text = ""), "null, ", "'" & EncryptString(txtPath.Text) & "', ")
    sSQL = sSQL & IIf((txtUserName.Text = ""), "null, ", "'" & EncryptString(txtUserName.Text) & "', ")
    sSQL = sSQL & IIf((txtPassword.Text = ""), "null, ", "'" & EncryptString(txtPassword.Text) & "', ")
    sSQL = sSQL & cboLoginType.ListIndex & ", "
    sSQL = sSQL & IIf((txtName.Text = ""), "null, ", "'" & ReplaceQuotes(txtName.Text) & "'")
    sSQL = sSQL & ")"
    oCon.Execute sSQL
  
    'close db connection
    oCon.Close
    Set oCon = Nothing
    HourglassOff
    Unload Me
    Exit Sub
    
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "frmActiveDirectoryNewServer.cmdOK_Click"
End Sub

'---------------------------------------------------------------------
Private Function NextConnectionOrder()
'---------------------------------------------------------------------
' ic 06/12/2005
' return next connectorder sequence
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsADServers As ADODB.Recordset
Dim vADServers As Variant

    sSQL = "SELECT MAX(CONNECTORDER) AS MAXCO FROM ACTIVEDIRECTORYSERVERS"
    Set rsADServers = New ADODB.Recordset
    rsADServers.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsADServers.RecordCount > 0 Then
        If (IsNull(rsADServers!MAXCO)) Then
            NextConnectionOrder = 1
        Else
            NextConnectionOrder = rsADServers!MAXCO + 1
        End If
    Else
        NextConnectionOrder = 1
    End If
    
    rsADServers.Close
    Set rsADServers = Nothing
    
End Function

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------
' ic 07/12/2005
' form load
'---------------------------------------------------------------------
    Call cboLoginType.AddItem("Current domain with logged in users credentials")
    Call cboLoginType.AddItem("Given domain with logged in users credentials")
    'Call cboLoginType.AddItem("Given domain with stored username and password")
    'Call cboLoginType.AddItem("Given domain with entered username and password")
    cboLoginType.ListIndex = 0
End Sub
