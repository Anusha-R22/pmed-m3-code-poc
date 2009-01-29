VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDatabaseTimezone 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Timezone"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1290
      TabIndex        =   3
      Top             =   1890
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   2475
      TabIndex        =   2
      Top             =   1890
      Width           =   1125
   End
   Begin MSComCtl2.UpDown spnTimezone 
      Height          =   345
      Left            =   3375
      TabIndex        =   0
      Top             =   1440
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   609
      _Version        =   393216
      BuddyControl    =   "txtTimezone"
      BuddyDispid     =   196610
      OrigLeft        =   1200
      OrigTop         =   480
      OrigRight       =   1425
      OrigBottom      =   1095
      Increment       =   60
      Max             =   720
      Min             =   -720
      SyncBuddy       =   -1  'True
      Wrap            =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtTimezone 
      Height          =   345
      Left            =   2475
      MaxLength       =   5
      TabIndex        =   1
      Top             =   1440
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   $"frmDatabaseTimezone.frx":0000
      Height          =   855
      Left            =   105
      TabIndex        =   5
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Use this dialog to specify the timezone where your database resides. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   105
      TabIndex        =   4
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmDatabaseTimezone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2003. All Rights Reserved
'   File:       frmDatabaseTimezone.frm
'   Author:     Matthew Martin, 10/06/2003
'   Purpose:    Allow user to view and modify database timezone.
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------

Option Explicit

Private moConnection As ADODB.Connection
Private mbInsert As Boolean

'-------------------------------------------------------------------------------
Public Sub Display(sDatabaseCode As String)
'-------------------------------------------------------------------------------
' MLM 11/06/03: Retreive the database timezone from the database and display it.
'-------------------------------------------------------------------------------

Dim oDatabase As MACROUserBS30.Database
Dim rsTimezone As ADODB.Recordset
Dim sMessage As String

    On Error GoTo ErrorHandler

    'Use User object to open a connection to the specified database
    Set oDatabase = New MACROUserBS30.Database
    If oDatabase.Load(SecurityADODBConnection, goUser.UserName, sDatabaseCode, "", False, sMessage) Then
        Set moConnection = New ADODB.Connection
        moConnection.Open oDatabase.ConnectionString
        Me.Icon = frmMenu.Icon
        Me.Caption = sDatabaseCode & " Timezone"
        Set rsTimezone = New ADODB.Recordset
        With rsTimezone
            .Open "SELECT SettingValue FROM MACRODBSetting WHERE SettingSection = 'timezone' AND SettingKey = 'dbtz'", _
                moConnection, adOpenStatic, adLockReadOnly, adCmdText
            mbInsert = .EOF
            If mbInsert Then
                txtTimezone.Text = "0"
            Else
                txtTimezone.Text = .Fields(0).Value
            End If
            .Close
        End With
        Me.Show vbModal
    Else
        DialogError sMessage
    End If
    
    Exit Sub
    
ErrorHandler:
    If MACROErrorHandler("frmDatabaseTimezone", Err.Number, Err.Description, "Display", Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'-------------------------------------------------------------------------------
Private Sub cmdCancel_Click()
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------

    Unload Me

End Sub

'-------------------------------------------------------------------------------
Private Sub cmdOK_Click()
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------

Dim sSQL As String

    On Error GoTo ErrorHandler

    'store the currently displayed timezone in db
    If mbInsert Then
        sSQL = "INSERT INTO MACRODBSetting (SettingSection, SettingKey, SettingValue)" & _
            " VALUES ('timezone', 'dbtz', '" & txtTimezone.Text & "')"
    Else
        sSQL = "UPDATE MACRODBSetting SET SettingValue = '" & txtTimezone.Text & _
            "' WHERE SettingSection = 'timezone' AND SettingKey = 'dbtz'"
    End If
    
    moConnection.Execute sSQL
    
    Unload Me
    
    Exit Sub
    
ErrorHandler:
    If MACROErrorHandler("frmDatabaseTimezone", Err.Number, Err.Description, "cmdOK_Click", Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'-------------------------------------------------------------------------------
Private Sub Form_Activate()
'-------------------------------------------------------------------------------
' MLM 11/06/03: Make the form user-friendly by automatically selecting the existing timezone
'-------------------------------------------------------------------------------
        
    txtTimezone.SetFocus

End Sub

'-------------------------------------------------------------------------------
Private Sub Form_Terminate()
'-------------------------------------------------------------------------------
'MLM 11/06/02: Tidy up.
'-------------------------------------------------------------------------------
    
    If Not moConnection Is Nothing Then
        moConnection.Close
        Set moConnection = Nothing
    End If
    
End Sub

'-------------------------------------------------------------------------------
Private Sub spnTimezone_Change()
'-------------------------------------------------------------------------------
' MLM 11/06/03: Clicking on the up and down arrows should select the timezone,
'   so that the user may replace it by typing if desired.
'-------------------------------------------------------------------------------
        
        txtTimezone.SetFocus

End Sub

'-------------------------------------------------------------------------------
Private Sub txtTimezone_Change()
'-------------------------------------------------------------------------------
' MLM 11/06/03: Validate the current text, and enable the OK button if it is a valid timezone.
'-------------------------------------------------------------------------------
    
    On Error GoTo ErrorHandler
    
    With txtTimezone
        If IsNumeric(txtTimezone.Text) Then
            'NB 720 mins = 12 hours
            cmdOK.Enabled = (Abs(CInt(.Text)) <= 720) And _
                InStr(.Text, ".") = 0 And InStr(.Text, ",") = 0
        Else
            cmdOK.Enabled = False
        End If
    End With
    
    Exit Sub
    
ErrorHandler:
    If MACROErrorHandler("frmDatabaseTimezone", Err.Number, Err.Description, "txtTimezone_Change", Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'-------------------------------------------------------------------------------
Private Sub txtTimezone_GotFocus()
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------

    With txtTimezone
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub
