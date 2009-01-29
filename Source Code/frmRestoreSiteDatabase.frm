VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRestoreSiteDatabase 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Restore Site Database"
   ClientHeight    =   4245
   ClientLeft      =   6870
   ClientTop       =   4725
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6540
      TabIndex        =   9
      Top             =   3780
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Create new or select existing database"
      Height          =   1335
      Left            =   60
      TabIndex        =   6
      Top             =   2340
      Width           =   7815
      Begin VB.CommandButton cmdRestore 
         Caption         =   "&Restore"
         Height          =   375
         Left            =   6480
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtDatabase 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   7575
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Select..."
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         ToolTipText     =   "Select an existing site database"
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdDatabase 
         Caption         =   "Crea&te..."
         Height          =   375
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Create new site database"
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.ListBox lstTrial 
      Height          =   1815
      Left            =   4140
      TabIndex        =   1
      Top             =   300
      Width           =   3615
   End
   Begin VB.ListBox lstSite 
      Height          =   1815
      Left            =   180
      TabIndex        =   0
      Top             =   300
      Width           =   3615
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sites (remote)"
      Height          =   2175
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   3855
   End
   Begin VB.Frame Frame3 
      Caption         =   "Studies"
      Height          =   2175
      Left            =   4020
      TabIndex        =   8
      Top             =   60
      Width           =   3855
   End
   Begin MSComDlg.CommonDialog dlgSaveAs 
      Left            =   5400
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRestoreSiteDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999. All Rights Reserved
'   File:       frmRestoreSiteDatabase.frm
'   Author:     Will Casey, March 2000
'   Purpose:    To enable a remote database to be recreated from the central database
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   TA 13/12/2000: Tidyied up dialog messages and titles
'   ASH 08/02/2002 Modified cmdDatabase_Click to fix bug 2.2.5 no.2
'   ZA 29/05/2002  Modified cmdDatabase_Click,cmdBrowse_Click to fix bug 16 in build 2.2.14
'   NCJ 15 Jan 04 - Corrected "Secuirty" typos
'   MLM 02/06/2008: Issue 3038: Order list of sites
'------------------------------------------------------------------------------------'


Option Explicit

Private mlRestoreTrialId As Long
Private msRestoreTrialSite As String
Private msRestoreTrialName As String
Private msRestoreDBPath As String
Private msRestoreDBPwd As String
Private msRestoreDBName As String
Private msDataSource As String
Private msConnection As String

Private mbSecurityDataRestored As Boolean

'------------------------------------------------------------------------------------'
Public Property Get RestoreDBName() As String
'------------------------------------------------------------------------------------'
'Name of the database to be restored
'------------------------------------------------------------------------------------'

    RestoreDBName = msRestoreDBName
    
End Property

'------------------------------------------------------------------------------------'
Public Property Get RestoreDataSource() As String
'------------------------------------------------------------------------------------'
'Data source of database to be restored
'------------------------------------------------------------------------------------'
    
    RestoreDataSource = msDataSource
    
End Property

'------------------------------------------------------------------------------------'
Public Property Get RestoreTrialId() As Long
'------------------------------------------------------------------------------------'
' The TrialId of the study to be restored
'------------------------------------------------------------------------------------'
    RestoreTrialId = mlRestoreTrialId
    
End Property

'------------------------------------------------------------------------------------'
Public Property Let RestoreTrialId(lRestoreTrialId As Long)
'------------------------------------------------------------------------------------'
' The TrialId of the study to be restored
'------------------------------------------------------------------------------------'
    mlRestoreTrialId = lRestoreTrialId
    
End Property

'------------------------------------------------------------------------------------'
Public Property Get RestoreTrialSite() As String
'------------------------------------------------------------------------------------'
' The TrialSite of the study to be restored
'------------------------------------------------------------------------------------'
    RestoreTrialSite = msRestoreTrialSite
    
End Property

'------------------------------------------------------------------------------------'
Public Property Let RestoreTrialSite(sRestoreTrialSite As String)
'------------------------------------------------------------------------------------'
' The TrialSite of the study to be restored
'------------------------------------------------------------------------------------'
    msRestoreTrialSite = sRestoreTrialSite
    
End Property

'------------------------------------------------------------------------------------'
Public Property Get RestoreTrialName() As String
'------------------------------------------------------------------------------------'
' The TrialName of the study to be restored
'------------------------------------------------------------------------------------'
    RestoreTrialName = msRestoreTrialName
    
End Property

'------------------------------------------------------------------------------------'
Public Property Let RestoreTrialName(sRestoreTrialName As String)
'------------------------------------------------------------------------------------'
' The TrialName of the study to be restored
'------------------------------------------------------------------------------------'
    msRestoreTrialName = sRestoreTrialName
    
End Property
'------------------------------------------------------------------------------------'
Public Property Get RestoreDBPwd() As String
'------------------------------------------------------------------------------------'
' The password for a target db
'------------------------------------------------------------------------------------'
    RestoreDBPwd = msRestoreDBPwd
    
End Property
'------------------------------------------------------------------------------------'
Public Property Get RestoreDBPath() As String
'------------------------------------------------------------------------------------'
' The path for a target db
'------------------------------------------------------------------------------------'
    RestoreDBPath = msRestoreDBPath
    
End Property

'----------------------------------------------------------------------------------------------
Private Sub cmdBrowse_Click()
'----------------------------------------------------------------------------------------------
' Browse to a database to restore to
'----------------------------------------------------------------------------------------------
' REVISIONS
' DPH 12/12/2002 - Changed to allow user to enter macro db password
' REM 12/02/2003 - Changed to handle SQL server and Oracle and not Access
'----------------------------------------------------------------------------------------------
Dim lSecAndMACRO As Long
Dim sMSG As String

    On Error GoTo ErrorHandler
    
    lSecAndMACRO = frmOptionMsgBox.Display(GetApplicationTitle, "Restore site database", "Please select one of the following:", "Security/MACRO database|MACRO database|Exit", "&OK", "", True, False)
    
    'exit sub if user clicked exit on optionbox
    If lSecAndMACRO = RestoreSite.ExitRestore Then Exit Sub
    
    'get connection string
    msConnection = CreateOrRegisterSecurityDB(True, True, True)
    
    'if only a MACRO DB then don't want to try and restore security database
    If lSecAndMACRO = RestoreSite.MACROOnly Then
        mbSecurityDataRestored = True
    ElseIf lSecAndMACRO = RestoreSite.SecurityAndMACRO Then
        mbSecurityDataRestored = False
    End If
    
    If msConnection = "" Then
        Call DialogInformation("Unable to find Database")
    Else
        'check if is a valid database
        If IsMACRODatabase(msConnection, lSecAndMACRO) Then
            Call SetDatabasePathText
            Call DialogInformation("Security and MACRO databases path set", "Restore site database")
        Else 'if not display correct message
            If lSecAndMACRO = RestoreSite.MACROOnly Then
                sMSG = "This is not a valid MACRO database"
            ElseIf lSecAndMACRO = RestoreSite.SecurityAndMACRO Then
                sMSG = "This is not a valid Security/MACRO database"
            End If
            
            Call DialogInformation(sMSG, "Restore site database")
            'disable the restore button
            cmdRestore.Enabled = False
        End If
        
    End If
    
Exit Sub
ErrorHandler:
         If Err.Number = cdlCancel Then ' The user clicked the cancel button on the common dialogue
            Exit Sub
         Else
           Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdBrowse_Click")
             Case OnErrorAction.Ignore
                 Resume Next
             Case OnErrorAction.Retry
                 Resume
             Case OnErrorAction.QuitMACRO
                 Call ExitMACRO
                 Call MACROEnd
            End Select
        End If
End Sub

'----------------------------------------------------------------------------------------------
Private Sub cmdDatabase_Click()
'----------------------------------------------------------------------------------------------
' Browse to a location to put the database in and then create the database using frmNewDatabase.NewMACRODatabaseAccess
'REM 31/01/03 - Changed routine to handle SQL Server and Oracle site databases as Access is no longer used
'----------------------------------------------------------------------------------------------

    On Error GoTo ErrorHandler

    'creates a seurity and MACRO database in one Schema
    msConnection = CreateNewSecurityDatabase(True)
    
    mbSecurityDataRestored = False
    
    If msConnection = "" Then
        Call DialogInformation("Unable to create new database")
    Else
    
        Call SetDatabasePathText
        
        Call DialogInformation("Security and MACRO databases created", "Restore site database")
        
    End If
    
Exit Sub
ErrorHandler:
         If Err.Number = cdlCancel Then ' The user clicked the cancel button on the common dialogue
            Exit Sub
         Else
           Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdDatabase_Click")
             Case OnErrorAction.Ignore
                 Resume Next
             Case OnErrorAction.Retry
                 Resume
             Case OnErrorAction.QuitMACRO
                 Call ExitMACRO
                 Call MACROEnd
            End Select
        End If
End Sub

'----------------------------------------------------------------------------------------------
Private Sub SetDatabasePathText()
'----------------------------------------------------------------------------------------------
'
'
'----------------------------------------------------------------------------------------------
Dim tCon As udtConnection
Dim sProvider As String
Dim sUserId As String
Dim sText As String

        tCon = Connection_AsType(msConnection)
        
        sProvider = tCon.Provider
        msDataSource = tCon.Datasource
        sUserId = tCon.UserId
        msRestoreDBName = tCon.Database
    
        If sProvider = CONNECTION_MSDAORA Then
            sText = "Provider = " & sProvider & "; " & "Data Source = " & msDataSource & "; " & "User ID = " & sUserId
        ElseIf sProvider = CONNECTION_SQLOLEDB Then
            sText = "Provider = " & sProvider & "; " & "Database = " & msRestoreDBName & "; " & "User ID = " & sUserId
        Else
            sText = "Unknown Provider: " & sProvider
        End If
        
        txtDatabase.Text = sText
        
        Call EnableRestore
End Sub


'----------------------------------------------------------------------------------------------
Private Sub cmdExit_Click()
'----------------------------------------------------------------------------------------------
' unload the form
'----------------------------------------------------------------------------------------------
     Unload Me

End Sub

'----------------------------------------------------------------------------------------------
Private Sub RefreshTrialsSites()
'----------------------------------------------------------------------------------------------
' Add all the Available trials and trial Sites to the lists when the form is
' called.
'----------------------------------------------------------------------------------------------

Dim rsTrials As ADODB.Recordset
Dim rsSites As ADODB.Recordset
Dim sSQL As String
Dim rsTrialName As ADODB.Recordset
Dim sTrialName As String
Dim sTrialSite As String
Dim lTrialId As Long

    On Error GoTo ErrHandler

    lstSite.Clear
    lstTrial.Clear
    
    'MLM 02/06/2008: Issue 3038: Order list of sites
    sSQL = "SELECT DISTINCT TrialSite.TrialSite FROM TrialSite, Site" _
        & " WHERE TrialSite.TrialSite = Site.Site" _
        & " AND Site.SiteLocation = 1" _
        & " ORDER BY TrialSite.TrialSite"
        
    Set rsSites = New ADODB.Recordset
    rsSites.Open sSQL, MacroADODBConnection, adOpenForwardOnly, , adCmdText
    
    With rsSites
        Do Until .EOF = True
            sTrialSite = rsSites!TrialSite
            lstSite.AddItem (sTrialSite)
             .MoveNext
        Loop
    End With
    If rsSites.RecordCount > 0 Then
        lstSite.Selected(0) = True
        Call lstSite_Click
    End If
        
    rsSites.Close
    Set rsSites = Nothing
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "RefreshTrialsSites")
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
Private Sub cmdRestore_Click()
'----------------------------------------------------------------------------------------------
' Show the Save as dialogue box Create the other connection object
'----------------------------------------------------------------------------------------------
Dim rsTrialID As ADODB.Recordset
Dim sSQL As String
Dim sMSG As String

    
    On Error GoTo ErrHandler
    
    'Get the study name from the Trial Id
    sSQL = " SELECT ClinicalTrialId from ClinicalTrial WHERE ClinicalTrialName = '" & msRestoreTrialName & "'"
    Set rsTrialID = New ADODB.Recordset
    rsTrialID.Open sSQL, MacroADODBConnection, adOpenKeyset, , adCmdText
    mlRestoreTrialId = rsTrialID!ClinicalTrialId
    Set rsTrialID = Nothing
    
    'Create the  connection object for the target database
    'Call InitializeCopyDataAdodbConnection(msRestoreDBPath)
    ' Start the data transfer
    Call DoDataTransfer(msConnection, mbSecurityDataRestored)

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdRestore_Click")
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
Private Sub Form_Load()
'----------------------------------------------------------------------------------------------
    
'----------------------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
                
    Me.Icon = frmMenu.Icon
    
    ' Set the path to nothing on every form load
    'Call ReSetDBPath
    
    txtDatabase.Enabled = False
    RefreshTrialsSites
    'disable the update button until the user chooses a TrialSite
    cmdRestore.Enabled = False
    
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

'----------------------------------------------------------------------------------------------
Private Sub lstSite_Click()
'----------------------------------------------------------------------------------------------
'When a site is clicked then show all the trials that have been set up at that site.
'----------------------------------------------------------------------------------------------
Dim rsTrialSite As ADODB.Recordset
Dim rsTrialName As ADODB.Recordset
Dim sSQL As String
    
    On Error GoTo ErrHandler
    
    lstTrial.Clear
    msRestoreTrialSite = Trim(lstSite.Text)
    sSQL = " SELECT ClinicalTrialId FROM TrialSite WHERE TrialSite = '" & msRestoreTrialSite & "'"
    Set rsTrialSite = New ADODB.Recordset
    rsTrialSite.Open sSQL, MacroADODBConnection, adOpenForwardOnly, , adCmdText
    
    With rsTrialSite
        Do Until .EOF = True
             mlRestoreTrialId = rsTrialSite!ClinicalTrialId
             sSQL = " SELECT ClinicalTrialName from ClinicalTrial WHERE ClinicalTrialId = " & mlRestoreTrialId
             Set rsTrialName = New ADODB.Recordset
             rsTrialName.Open sSQL, MacroADODBConnection, adOpenKeyset, , adCmdText
             lstTrial.AddItem (rsTrialName!ClinicalTrialName)
            .MoveNext
        Loop
    End With
                
    Set rsTrialName = Nothing
    Set rsTrialSite = Nothing


   Call EnableRestore
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "lstSite_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'---------------------------------------------------------------------
Private Sub lstTrial_Click()
'---------------------------------------------------------------------
'enable the restore button if something has been selected in both lists.
'Store the trial name in the msRestoreTrialName variable
'---------------------------------------------------------------------
    On Error GoTo ErrHandler
            
    msRestoreTrialName = lstTrial.Text
    
    Call EnableRestore
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, " lstTrial_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub


'---------------------------------------------------------------------
Private Sub EnableRestore()
'---------------------------------------------------------------------
'If something has been chosen in both listboxes and one of the option
' buttons then enable the restore button
'---------------------------------------------------------------------

    If lstSite.ListIndex = -1 Or lstTrial.ListIndex = -1 Or msConnection = "" Then
        cmdRestore.Enabled = False
    Else
        cmdRestore.Enabled = True
    End If
    
End Sub

'---------------------------------------------------------------------
Private Function IsAMacroDatabase(msRestoreDBPath As String, sDatabasePassword As String) As Boolean
'---------------------------------------------------------------------
' Check to see if the item being browsed to is a valid Macro database
'---------------------------------------------------------------------

Dim sSQL As String
Dim ADODBConnection As ADODB.Connection
Dim rsTest As ADODB.Recordset

   'Assume its not a Macro database
   IsAMacroDatabase = False

   ' set up handler for if we cant open the chosen item
    On Error GoTo ErrCantOpenDB:

    'Try and open the chosen item
    Set ADODBConnection = New ADODB.Connection
    'ZA 07/06/2002 - changed optional password parameter to 5th input parameter rather than 3rd
    'parameter as it will cause failure to open the MACRO database
    ADODBConnection.Open Connection_String(CONNECTION_MSJET_OLEDB_40, msRestoreDBPath, , , sDatabasePassword)


    On Error GoTo ErrHandler

    ' Test for MACRO DB by looking for ClinicalTrialID = 0
    sSQL = "SELECT ClinicalTrialId FROM ClinicalTrial WHERE ClinicalTrialId = 0"
    Set rsTest = New ADODB.Recordset
    rsTest.Open sSQL, ADODBConnection, adOpenKeyset, , adCmdText
    If rsTest.RecordCount > 0 Then
        ' Seems OK
        rsTest.Close
        Set rsTest = Nothing
        ADODBConnection.Close
        Set ADODBConnection = Nothing
        IsAMacroDatabase = True
    Else
        IsAMacroDatabase = False
    End If



Exit Function

ErrCantOpenDB:
    Set rsTest = Nothing
    Set ADODBConnection = Nothing
    MsgBox "The item you are trying to open is not a MACRO database.", vbOKOnly, "MACRO"
    Exit Function

ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, " lstTrial_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Function
'
''---------------------------------------------------------------------
'Private Sub ReSetDBPath()
''---------------------------------------------------------------------
'' Set DBPath to nothing
''---------------------------------------------------------------------
'
'    msRestoreDBPath = vbNullString
'
'End Sub
