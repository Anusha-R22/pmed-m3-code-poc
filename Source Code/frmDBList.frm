VERSION 5.00
Begin VB.Form frmDBList 
   Caption         =   "Form1"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2520
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   2520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   1320
      TabIndex        =   2
      Top             =   1980
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   345
      Left            =   60
      TabIndex        =   1
      Top             =   1980
      Width           =   1125
   End
   Begin VB.ListBox lstDatabases 
      Height          =   1815
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2355
   End
End
Attribute VB_Name = "frmDBList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mconSecurity As ADODB.Connection
Private msUserName As String
Private msSiteCode As String
Private msHTTPAddress As String
Private msIISUserName As String
Private msIISPassword As String
Private mnPortNumber As Integer
Private msDatabaseCode As String
Private msErrMessage As String
Private mbOKClicked As Boolean
Private mconMACRO As ADODB.Connection

'---------------------------------------------------------------------
Public Function Display(sSecurityCon As String, sUsername As String, ByRef sSiteCode As String, _
                        ByRef sHTTPAddress As String, ByRef sIISUserName As String, ByRef sIISPassword As String, _
                        ByRef nPortNumber As Integer, ByRef sErrmessage As String) As String
'---------------------------------------------------------------------
'REM 06/12/02
'Displays the form with all users databases and returns the selected databases, Sitecode, HTTPAddress,
' IIS Username and Password, Port Number and any error message
'---------------------------------------------------------------------
Dim vDatabases As Variant
Dim nCount As Integer

    'create and open a new security connection
    Set mconSecurity = New ADODB.Connection
    
    FormCentre Me
    Me.Icon = frmMenu.Icon

    msUserName = sUsername

    mbOKClicked = False

    'setup security connection
    mconSecurity.Open sSecurityCon
    mconSecurity.CursorLocation = adUseClient
    
    vDatabases = GetUserDatabases(nCount)
    
    Select Case nCount
    Case 0
        sSiteCode = ""
        Display = ""
        sIISUserName = ""
        sIISPassword = ""
        sHTTPAddress = ""
        nPortNumber = -1
        sErrmessage = msErrMessage
        Exit Function
    
    Case 1 'if there is one DB then use it
        Call GetSettings(vDatabases(0, 0))
        Display = msDatabaseCode
        sSiteCode = msSiteCode
        sHTTPAddress = msHTTPAddress
        sIISUserName = msIISUserName
        sIISPassword = msIISPassword
        nPortNumber = mnPortNumber
        sErrmessage = msErrMessage
        Exit Function
        
    Case Else 'let the user choose a database
    
        Call LoadDatabaseListBox(vDatabases)
        Me.Show vbModal
       
    End Select
    
    If mbOKClicked Then
        Display = msDatabaseCode
        sSiteCode = msSiteCode
        sHTTPAddress = msHTTPAddress
        sIISUserName = msIISUserName
        sIISPassword = msIISPassword
        nPortNumber = mnPortNumber
        sErrmessage = msErrMessage
    Else
        sSiteCode = ""
        Display = ""
        sIISUserName = ""
        sIISPassword = ""
        sHTTPAddress = ""
        nPortNumber = -1
        sErrmessage = "You have not chosen a database, MACRO will now shut down"
    End If
    
End Function

'---------------------------------------------------------------------
Private Function GetUserDatabases(ByRef nCount As Integer) As Variant
'---------------------------------------------------------------------
'REM 06/12/02
'Loads the database list box with all the users databases
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsDBs As ADODB.Recordset
Dim vDatabases As Variant
Dim i As Integer
Dim sDatabaseCode As String

    
    'get list of user databases
    sSQL = "SELECT DatabaseCode FROM UserDatabase" _
        & " WHERE UserName = '" & msUserName & "'"
    Set rsDBs = New ADODB.Recordset
    rsDBs.Open sSQL, mconSecurity, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsDBs.RecordCount <> 0 Then
        nCount = rsDBs.RecordCount
        vDatabases = rsDBs.GetRows
    Else
        nCount = 0
        vDatabases = Null
    End If
    
    GetUserDatabases = vDatabases
    
    rsDBs.Close
    Set rsDBs = Nothing
    
End Function

'---------------------------------------------------------------------
Private Sub LoadDatabaseListBox(vDatabases As Variant)
'---------------------------------------------------------------------
'REM 09/12/02
'Loads the list box with all user databases
'---------------------------------------------------------------------
Dim i As Integer
Dim sDatabaseCode As String

    If Not IsNull(vDatabases) Then
        For i = 0 To UBound(vDatabases, 2)
            sDatabaseCode = vDatabases(0, i)
            lstDatabases.AddItem sDatabaseCode
        Next
        
        'select the first one in the list
        lstDatabases.Selected(0) = True
    End If
    
End Sub

Private Sub cmdCancel_Click()
    mbOKClicked = False
End Sub

'---------------------------------------------------------------------
Private Sub cmdOK_Click()
'---------------------------------------------------------------------
'REM 06/12/02
'
'---------------------------------------------------------------------

    mbOKClicked = True

    Call GetSettings(lstDatabases.Text)
    
    Unload Me
    
End Sub

'---------------------------------------------------------------------
Private Sub GetSettings(ByVal sDatabaseCode As String)
'---------------------------------------------------------------------
'REM 09/12/02
'Asigns all the setting to modular level variables
'---------------------------------------------------------------------
Dim oDatabase As MACROUserBS30.Database
Dim sMsg As String
Dim sConnection As String

    msDatabaseCode = sDatabaseCode
    
    Set oDatabase = New MACROUserBS30.Database
    
    Call oDatabase.Load(mconSecurity, msUserName, msDatabaseCode, "", False, sMsg)
    sConnection = oDatabase.ConnectionString
    
    'create the connection
    Set mconMACRO = New ADODB.Connection
    mconMACRO.Open sConnection
    mconMACRO.CursorLocation = adUseClient
    
    msSiteCode = GetMacroDBSetting("datatransfer", "dbsitename", mconMACRO)
    If msSiteCode = "" Then
        msErrMessage = "This is a server database, please contact your system administrator if you have forgotten your password"
    End If
    
    Call GetHTTPAddressAndPortNumber

End Sub

'---------------------------------------------------------------------
Private Sub GetHTTPAddressAndPortNumber()
'---------------------------------------------------------------------
'REM 06/12/02
'Returns the most current HTTP address and port number from the trial office table
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsAddress As ADODB.Recordset
Dim vAddress As Variant

    sSQL = "SELECT * FROM TrialOffice" _
        & " ORDER BY EffectiveFrom DESC"
    Set rsAddress = New ADODB.Recordset
    rsAddress.Open sSQL, mconMACRO, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsAddress.RecordCount > 0 Then
        vAddress = rsAddress.GetRows
        msHTTPAddress = vAddress(3, 0)
        msIISUserName = vAddress(6, 0)
        msIISPassword = vAddress(7, 0)
        mnPortNumber = vAddress(8, 0)
    End If

End Sub
