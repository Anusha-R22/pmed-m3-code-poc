VERSION 5.00
Begin VB.Form frmConnectionString 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connection String"
   ClientHeight    =   3255
   ClientLeft      =   6480
   ClientTop       =   4275
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4860
      TabIndex        =   13
      Top             =   2820
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3540
      TabIndex        =   12
      Top             =   2820
      Width           =   1215
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "&Test"
      Height          =   375
      Left            =   60
      TabIndex        =   11
      Top             =   2820
      Width           =   1215
   End
   Begin VB.Frame fraMain 
      Height          =   2715
      Left            =   60
      TabIndex        =   14
      Top             =   0
      Width           =   6015
      Begin VB.Frame Frame2 
         Caption         =   "Server/Site"
         Height          =   975
         Left            =   2400
         TabIndex        =   22
         Top             =   180
         Width           =   3495
         Begin VB.OptionButton optServer 
            Caption         =   "Server"
            Height          =   345
            Left            =   120
            TabIndex        =   2
            Top             =   230
            Width           =   800
         End
         Begin VB.OptionButton optSite 
            Caption         =   "Site"
            Height          =   345
            Left            =   120
            TabIndex        =   3
            Top             =   550
            Width           =   615
         End
         Begin VB.TextBox txtSiteName 
            Height          =   300
            Left            =   1660
            MaxLength       =   8
            TabIndex        =   4
            Top             =   540
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Site Name"
            Height          =   315
            Left            =   840
            TabIndex        =   23
            Top             =   620
            Width           =   795
         End
      End
      Begin VB.TextBox txtTNS 
         Height          =   300
         Left            =   1500
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1215
         Width           =   1695
      End
      Begin VB.Frame fraDatabaseoptions 
         Caption         =   "Database Options"
         Height          =   975
         Left            =   120
         TabIndex        =   20
         Top             =   180
         Width           =   2235
         Begin VB.OptionButton optSQL 
            Caption         =   "SQL Server / MSDE"
            Height          =   345
            Left            =   120
            TabIndex        =   1
            Top             =   550
            Width           =   1755
         End
         Begin VB.OptionButton optORACLE 
            Caption         =   "Oracle"
            Height          =   345
            Left            =   120
            TabIndex        =   0
            Top             =   230
            Width           =   855
         End
      End
      Begin VB.TextBox txtAlias 
         Height          =   300
         Left            =   1500
         MaxLength       =   15
         TabIndex        =   10
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox txtServer 
         Height          =   300
         Left            =   4065
         TabIndex        =   6
         Top             =   1575
         Width           =   1695
      End
      Begin VB.TextBox txtDatabase 
         Height          =   300
         Left            =   4065
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1215
         Width           =   1695
      End
      Begin VB.TextBox txtPassword 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1500
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox txtUID 
         Height          =   300
         Left            =   1500
         TabIndex        =   8
         Top             =   1575
         Width           =   1695
      End
      Begin VB.Label lblTns 
         Caption         =   "TNS Name"
         Height          =   195
         Left            =   135
         TabIndex        =   21
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label lblAlias 
         Caption         =   "&MACRO DB Alias"
         Height          =   255
         Left            =   135
         TabIndex        =   19
         Top             =   2340
         Width           =   1275
      End
      Begin VB.Label lblServer 
         Caption         =   "Server"
         Height          =   255
         Left            =   3300
         TabIndex        =   18
         Top             =   1620
         Width           =   555
      End
      Begin VB.Label lblDatabase 
         Caption         =   "Database"
         Height          =   255
         Left            =   3300
         TabIndex        =   17
         Top             =   1260
         Width           =   795
      End
      Begin VB.Label lblPassword 
         Caption         =   "Password"
         Height          =   195
         Left            =   135
         TabIndex        =   16
         Top             =   1980
         Width           =   735
      End
      Begin VB.Label lblUID 
         Caption         =   "User Name"
         Height          =   255
         Left            =   135
         TabIndex        =   15
         Top             =   1620
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmConnectionString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2003-2004. All Rights Reserved
'   File:       frmConnectionString.frm
'   Author:     Ashitei Trebi-Ollennu, January 2003
'   Purpose:    Builds connection strings
'------------------------------------------------------------------------------
'REVISIONS:
'REM 10/02/04 - Added routine EnableTestButton
'-------------------------------------------------------------------------
    

Option Explicit
Private msConnectionstring As String
Private msSiteCode As String
Private msDBAlias As String
Private mbIsLoading As Boolean
Private mbRegister As Boolean
Private mbUserChangedAlias As Boolean


'-------------------------------------------------------------------------
Public Function Display(bRegister As Boolean, bDBAlias As Boolean, bSite As Boolean, Optional ByRef sDBAlias As String, _
                        Optional ByRef sSiteCode As String, Optional bMSDE As Boolean = False) As String
'-------------------------------------------------------------------------
'
'-------------------------------------------------------------------------
    mbIsLoading = True
    
    cmdOK.Enabled = False
    
    HourglassSuspend
    
    mbUserChangedAlias = False
    
    mbRegister = bRegister
    
    txtAlias.Enabled = bDBAlias
    optSite.Enabled = bSite
    optServer.Enabled = bSite
    optServer.Value = True
    txtSiteName.Enabled = False
    
    Call EnableTestButton(oracle80)
    
    optORACLE.Enabled = Not bMSDE
    If bMSDE Then
        optSQL.Value = bMSDE
    End If
    
    frmConnectionString.Caption = GetApplicationTitle
    
    mbIsLoading = False
    Me.Show vbModal
    
    HourglassResume
    
    sDBAlias = msDBAlias
    sSiteCode = msSiteCode
    Display = msConnectionstring
     
End Function

'----------------------------------------------------------------------
Private Sub cmdCancel_Click()
'----------------------------------------------------------------------
'
'----------------------------------------------------------------------

    msConnectionstring = vbNullString
    Unload Me

End Sub

'-----------------------------------------------------------------------
Private Sub cmdOK_Click()
'-----------------------------------------------------------------------
'
'-----------------------------------------------------------------------

    'REM 23/01/03 - added Server/Site settings
    If optSite.Value Then
        msSiteCode = txtSiteName.Text
        If msSiteCode = "" Then
            Call DialogWarning("Please enter a site name.")
            txtSiteName.SetFocus
            Exit Sub
        End If
    Else
        msSiteCode = ""
    End If
    
    msDBAlias = Trim(txtAlias.Text)
    
    Unload Me
    
    DoEvents

End Sub

'------------------------------------------------------------------------
Private Sub cmdTest_Click()
'------------------------------------------------------------------------
'
'------------------------------------------------------------------------

    TestConnectString
    
End Sub

'-------------------------------------------------------------------------
Private Sub EnableTestButton(nDatabaseType As MACRODatabaseType)
'-------------------------------------------------------------------------
'REM 10/02/04 - Don't enable Test button fro SQL Server unless there is something
' in the server and database fields
'-------------------------------------------------------------------------
    
    If nDatabaseType = oracle80 Then
        'Don't have to worry about disableing the Test button as it will take care of itself
        cmdTest.Enabled = True
    ElseIf nDatabaseType = sqlserver Then
        If (Trim(txtServer.Text) <> "") And (Trim(txtDatabase.Text) <> "") Then
            cmdTest.Enabled = True
        Else
            cmdTest.Enabled = False
        End If
    End If
     
End Sub

'-------------------------------------------------------------------------
Private Sub Form_Load()
'-------------------------------------------------------------------------
'
'-------------------------------------------------------------------------
    
    Me.Icon = frmMenu.Icon
    FormCentre Me
    optORACLE.Value = True
    lblServer.Visible = False
    txtServer.Visible = False
    lblDatabase.Visible = False
    txtDatabase.Visible = False

End Sub


'-------------------------------------------------------------------------
Private Sub optORACLE_Click()
'-------------------------------------------------------------------------
'
'-------------------------------------------------------------------------

    lblTns.Visible = True
    txtTNS.Visible = True
    lblServer.Visible = False
    txtServer.Visible = False
    lblDatabase.Visible = False
    txtDatabase.Visible = False
    txtTNS.Text = ""
    txtDatabase.Text = ""
    txtUID.Text = ""
    txtAlias.Text = ""
    txtPassword.Text = ""
    txtServer.Text = ""
    txtDatabase.Text = ""
    txtTNS.Width = 3000
    
    Call EnableTestButton(oracle80)

End Sub

'--------------------------------------------------------------------------
Private Sub optSQL_Click()
'--------------------------------------------------------------------------
'show SQL Server controls
'--------------------------------------------------------------------------
    
    lblServer.Top = lblTns.Top
    txtServer.Top = txtTNS.Top
    
    lblServer.Left = lblUID.Left
    txtServer.Left = txtTNS.Left
    
    lblDatabase.Top = lblTns.Top
    txtDatabase.Left = (txtServer.Width + lblDatabase.Left) - (lblDatabase.Width + 120)
    txtDatabase.Top = txtTNS.Top
        
    lblServer.Visible = True
    txtServer.Visible = True
    lblDatabase.Visible = True
    txtDatabase.Visible = True
    lblTns.Visible = False
    txtTNS.Visible = False
    
    txtTNS.Text = ""
    txtDatabase.Text = ""
    txtUID.Text = ""
    txtAlias.Text = ""
    txtPassword.Text = ""
    txtServer.Text = ""
    txtDatabase.Text = ""

    Call EnableTestButton(sqlserver)

End Sub

'------------------------------------------------------------------------------
Private Sub TestConnectString()
'------------------------------------------------------------------------------
Dim oCon As ADODB.Connection
Dim lConErrNo As Long
Dim sConErr As String
Dim sVersion As String

    On Error GoTo ErrHandler
    
    HourglassOn
    
    Set oCon = New ADODB.Connection
    If optORACLE Then
        msConnectionstring = Connection_String(CONNECTION_MSDAORA, Trim(txtTNS.Text), , _
            Trim(txtUID.Text), Trim(txtPassword.Text))
    Else
    
        msConnectionstring = Connection_String(CONNECTION_SQLOLEDB, Trim(txtServer.Text), _
        "", Trim(txtUID.Text), Trim(txtPassword.Text))
        On Error Resume Next
        oCon.Open msConnectionstring
        oCon.Execute "Create database " & Trim(txtDatabase.Text)
        oCon.Close
        
        On Error GoTo ErrHandler
    
        msConnectionstring = Connection_String(CONNECTION_SQLOLEDB, Trim(txtServer.Text), _
        Trim(txtDatabase.Text), Trim(txtUID.Text), Trim(txtPassword.Text))
    End If
   
    On Error Resume Next
    oCon.Open msConnectionstring
    sConErr = Err.Description
    lConErrNo = Err.Number
    Err.Clear
    
    On Error GoTo ErrHandler
    
    HourglassOff
    
    If lConErrNo <> 0 Then
        DialogError ("The connection to the specified database failed because of the following:") & vbCrLf _
        & sConErr
        cmdOK.Enabled = False
        Exit Sub
    Else
    'Check if the database already exists
        sVersion = ""
    
        If (CheckForMacroTables(oCon, sVersion) = True) And (mbRegister = False) Then
            'Database exists, (i.e. contains Macro tables), Do not Create
            DialogError ("The connection details you have provided refer to an already existing MACRO database (Version " & sVersion & ")." _
                & vbCr & "Database Creation cannot proceed.")
            cmdOK.Enabled = False
        Else
            DialogInformation ("The connection to the specified database tested successfully.")
            cmdOK.Enabled = True
        End If
    End If
      
    oCon.Close
    Set oCon = Nothing
 
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmConnectionString.TestConnectString"
End Sub

'-----------------------------------------------------------------
Private Sub txtAlias_Change()
'-----------------------------------------------------------------
'
'-----------------------------------------------------------------
Dim sCode As String
Dim iPos As Integer

    iPos = txtAlias.SelStart

    'If the form is loading then don't run the change event procedures
    If mbIsLoading = False Then
        
        sCode = txtAlias.Text
        ' check that the characters entered are valid
        If Not gblnValidString(sCode, valAlpha + valNumeric + valUnderscore + valDecimalPoint) Then
            txtAlias.Text = txtAlias.Tag
            txtAlias.SelStart = iPos
        End If

    End If

    txtAlias.Tag = Trim(txtAlias.Text)

End Sub

'------------------------------------------------------------------
Private Sub txtAlias_KeyPress(KeyAscii As Integer)
'------------------------------------------------------------------

    mbUserChangedAlias = True
    
End Sub

'------------------------------------------------------------------
Private Sub txtAlias_LostFocus()
'------------------------------------------------------------------
'
'------------------------------------------------------------------
    
    If Not gblnValidString(txtAlias.Text, valAlpha + valNumeric + valSpace + valUnderscore) Then
        Call DialogInformation("MACRO Alias contains invalid characters")
    End If

End Sub

'------------------------------------------------------------------
Private Sub txtDatabase_Change()
'------------------------------------------------------------------
'
'------------------------------------------------------------------

    cmdOK.Enabled = False

    If Not mbIsLoading Then
        If txtAlias.Text = vbNullString Or Not mbUserChangedAlias Then
            txtAlias.Text = txtDatabase.Text
        End If
        
        Call EnableTestButton(sqlserver)
    
    End If

End Sub

'-------------------------------------------------------------------
Private Sub txtDatabase_LostFocus()
'-------------------------------------------------------------------
'
'-------------------------------------------------------------------
    
    If Not gblnValidString(txtDatabase.Text, valAlpha + valNumeric + valSpace + valUnderscore) Then
        Call DialogInformation("Database name contains invalid characters")
    End If

End Sub

'--------------------------------------------------------------------
Private Sub txtPassword_Change()
'--------------------------------------------------------------------
'
'--------------------------------------------------------------------
    
    cmdOK.Enabled = False

End Sub

'---------------------------------------------------------------------
Private Sub txtPassword_LostFocus()
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------
    
    If Not gblnValidString(txtPassword.Text, valAlpha + valNumeric + valSpace + valUnderscore) Then
        Call DialogInformation("Password contains invalid characters")
    End If

End Sub

'----------------------------------------------------------------------
Private Sub txtServer_Change()
'----------------------------------------------------------------------
'
'----------------------------------------------------------------------
    
    cmdOK.Enabled = False
    
    Call EnableTestButton(sqlserver)

End Sub

'-----------------------------------------------------------------------
Private Sub txtServer_LostFocus()
'-----------------------------------------------------------------------
'
'-----------------------------------------------------------------------
    
    If Not gblnValidString(txtServer.Text, valAlpha + valNumeric + valSpace + valUnderscore + valDecimalPoint) Then
        Call DialogInformation("Server name contains invalid characters")
    End If


End Sub

'-----------------------------------------------------------------------
Private Sub txtSiteName_Change()
'-----------------------------------------------------------------------
'
'-----------------------------------------------------------------------
Dim sCode As String
Dim iPos As Integer

    iPos = txtSiteName.SelStart

    'If the form is loading then don't run the change event procedures
    If mbIsLoading = False Then
        'make the site name lower case only

        txtSiteName.Text = LCase(txtSiteName.Text)
        
        sCode = txtSiteName.Text
        ' check that the character entered is valid
        If Not gblnValidString(sCode, valAlpha + valNumeric) Then
            txtSiteName.Text = txtSiteName.Tag
            txtSiteName.SelStart = iPos
            
        End If
    
        ' check that the first char is not numeric
        If sCode <> vbNullString Then
            If gblnValidString(Left$(sCode, 1), valNumeric) Then
            txtSiteName.Text = txtSiteName.Tag
            txtSiteName.SelStart = iPos
                
            End If
        End If

    End If
    
    If iPos > 0 Then
        txtSiteName.SelStart = iPos
    End If

    txtSiteName.Tag = Trim(txtSiteName.Text)

End Sub

'-----------------------------------------------------------------------
Private Sub txtTNS_Change()
'-----------------------------------------------------------------------
'
'-----------------------------------------------------------------------

     cmdOK.Enabled = False

End Sub

'------------------------------------------------------------------------
Private Sub txtTNS_LostFocus()
'------------------------------------------------------------------------
'
'------------------------------------------------------------------------
    
    If Not gblnValidString(txtTNS.Text, valAlpha + valNumeric + valSpace + valUnderscore + valDecimalPoint) Then
        Call DialogInformation("TNS contains invalid characters")
    End If

End Sub

'-----------------------------------------------------------------------
Private Sub txtUID_Change()
'-----------------------------------------------------------------------
'
'-----------------------------------------------------------------------
    
    cmdOK.Enabled = False

End Sub

'------------------------------------------------------------------------
Private Sub txtUID_LostFocus()
'------------------------------------------------------------------------
'
'------------------------------------------------------------------------
    
    If Not gblnValidString(txtUID.Text, valAlpha + valNumeric + valSpace + valUnderscore + valDecimalPoint) Then
        Call DialogInformation("UID contains invalid characters")
    End If

End Sub

'--------------------------------------------------------------------
Private Sub optServer_Click()
'--------------------------------------------------------------------
    
    txtSiteName.Enabled = False
    
End Sub

'--------------------------------------------------------------------
Private Sub optSite_Click()
'--------------------------------------------------------------------
'REM 23/01/03
'When user clicks the Site option button must disable the OK button if there is no site name in the Site Name text box
'--------------------------------------------------------------------
    
    txtSiteName.Enabled = True
    
    If txtSiteName.Text = "" Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If

End Sub
