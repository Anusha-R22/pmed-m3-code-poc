VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUsers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Users"
   ClientHeight    =   4230
   ClientLeft      =   4485
   ClientTop       =   4515
   ClientWidth     =   6330
   Icon            =   "frmUsers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkEnabled 
      Caption         =   "Enabled"
      Height          =   375
      Left            =   4800
      TabIndex        =   11
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Width           =   5895
      Begin VB.CommandButton cmdApply 
         Caption         =   "&Apply"
         Height          =   495
         Left            =   2520
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Close"
         Height          =   495
         Left            =   4200
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdEditUserName 
      Caption         =   "&Edit user name"
      Height          =   495
      Left            =   4800
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdNewUser 
      Caption         =   "Add new &user"
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   720
      Width           =   3375
   End
   Begin VB.ComboBox cboUsers 
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin MSComctlLib.ListView lvwUserDatabase 
      Height          =   1815
      Left            =   1200
      TabIndex        =   5
      Top             =   1080
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   3201
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList Img1 
      Left            =   -240
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsers.frx":0442
            Key             =   "BoxNoTick"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsers.frx":0896
            Key             =   "BoxTick"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblUserCode 
      Caption         =   "User name:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Full real name:"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   840
      Width           =   1155
   End
   Begin VB.Label Label3 
      Caption         =   "User  databases:"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   855
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1999. All Rights Reserved
'   File:       frmUsers.frm
'   Author:     Will Casey 22/10/99
'   Purpose:    Maintain users and the databases which a user can access.
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
'   Revisions:
'WillC      11/10/99 Added the error handlers
'Mo Morris  22/12/99
'           in the same way as cmdNewUser, cmdEditUserName is now only enabled when
'           the user has functionAccess gsFnCreateNewUser. The check is made once in
'           form_load and its setting stored in mbUserHasFnCreateNewUser
'           WillC 26/1/00 SR 2771 Added change to allow multiple additions of databases at one
'           time.
'           WillC 21/2/00 Added msgbox to show which databases a user is assigned to
'           TA 08/05/2000   removed subclassing
'           WillC 1/8/00 Added  EnableApply() toCheck user rights before allowing the Apply button to be enabled.
'           WillC 7/8/00 SR2956   Changed labels
'           TA 12/12/2000: Changed msgbox text and used dialogs from library
'           ZA 18/09/2002: creates List_user.js file upon enabling/disabling a user
'--------------------------------------------------------------------------------

Option Explicit
Option Compare Binary
Option Base 0
' this collection holds the usernames and usercodes
Private suCodeCol As Collection


Private mbUserHasFnCreateNewUser As Boolean
'WillC 13/3/00 Added mb variable to handle a prorammatic click event.
Private mbProgrammaticCheckbox As Boolean

'----------------------------------------------------------------------------------------------
Private Sub chkEnabled_Click()
'----------------------------------------------------------------------------------------------
' SR3209 2985  This is the user clicking so its not a Programmatic check so mbProgrammaticCheckbox = false
'----------------------------------------------------------------------------------------------
    
 'Override any programatic info as this is a user action so toggle it from true to false...
 'mbProgrammaticCheckbox = False
 
    If mbProgrammaticCheckbox = True Then
        mbProgrammaticCheckbox = False
        Exit Sub
    Else
        Call EnableApply
        'cmdApply.Enabled = True
        mbProgrammaticCheckbox = False
    End If
    
End Sub

'----------------------------------------------------------------------------------------------
Private Sub cmdApply_Click()
'----------------------------------------------------------------------------------------------
' If a database is checked then add it to the userdatabase table
'----------------------------------------------------------------------------------------------
Dim sUserCode As String
Dim sDatabaseDescription As String
Dim n As Integer
Dim nEnabled As Integer
Dim sMsg As String
Dim sDBList As String

    On Error GoTo ErrHandler

    sUserCode = cboUsers.Text
    nEnabled = chkEnabled.Value
    
    If sUserCode = "" Then
       Call DialogError("Please select a user.")
       Exit Sub
    Else
    

    'WillC 26/1/00 SR 2771 Added change to allow multiple additions of databases at one
    ' time.
        For n = 1 To lvwUserDatabase.ListItems.Count
             If lvwUserDatabase.ListItems(n).SmallIcon = "BoxTick" Then
                sDatabaseDescription = lvwUserDatabase.ListItems(n).Text
              Call AddDatabaseToUser(sUserCode, sDatabaseDescription)
             ElseIf lvwUserDatabase.ListItems(n).SmallIcon = "BoxNoTick" Then
                sDatabaseDescription = lvwUserDatabase.ListItems(n).Text
               Call RemoveDatabaseFromUser(sUserCode, sDatabaseDescription)
             End If
        Next n
          Call gdsUserEnabled(sUserCode, nEnabled)
    End If
        
    'WillC 21/2/00 Added msgbox to show which databases a user is assigned to
    For n = 1 To lvwUserDatabase.ListItems.Count
        If lvwUserDatabase.ListItems(n).SmallIcon = "BoxTick" Then
           sDatabaseDescription = lvwUserDatabase.ListItems(n).Text
           sDBList = sDBList & " " & sDatabaseDescription & vbCrLf
        End If
    Next n
    
    'WillC 23/3/00 SR3292
    'create List_users.js file now
    If chkEnabled.Value = vbChecked Then    'Checked
        sMsg = " User " & txtUserName.Text & " has been assigned" & vbCrLf
        sMsg = sMsg & " to the following database(s). " & vbCrLf
        CreateUsersList
    ElseIf chkEnabled.Value = vbUnchecked Then 'Unchecked
        Call DialogInformation(txtUserName.Text & "'s user account is now disabled.")
        CreateUsersList
        cmdApply.Enabled = False
        Exit Sub
    End If
    
    If sDBList = "" Then
        Call DialogInformation(" User " & txtUserName.Text & " is not assigned to a database.")
    ElseIf sDBList = "" And chkEnabled = vbUnchecked Then
        Call DialogInformation(" User " & txtUserName.Text & " is not assigned to a database." & vbCrLf _
            & "The user has also been disabled.")
    Else
        Call DialogInformation(sMsg & sDBList)
    End If
    
    'WillC 23/2/2000
    cmdApply.Enabled = False
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdApply_Click")
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
' show the form to edit a user name
'----------------------------------------------------------------------------------------------
    Unload Me
End Sub

'----------------------------------------------------------------------------------------------
Private Sub cmdEditUserName_Click()
'----------------------------------------------------------------------------------------------
' show the form to edit a user name
'----------------------------------------------------------------------------------------------
     
     frmEditUserName.Show vbModal
    
End Sub

'----------------------------------------------------------------------------------------------
Private Sub cmdNewUser_Click()
'----------------------------------------------------------------------------------------------
' show the form to set up a new user
'----------------------------------------------------------------------------------------------
    
    frmNewUser.Show vbModal
End Sub

'----------------------------------------------------------------------------------------------
Private Sub Form_Load()
'----------------------------------------------------------------------------------------------
'Open the Database, and Load all the Lists with the data, add the columns to the
' listview, Disable any buttons that could cause a problem at this stage.
'----------------------------------------------------------------------------------------------
Dim clmX As MSComctlLib.ColumnHeader
Dim n As Integer

    On Error GoTo ErrHandler

    If goUser.CheckPermission(gsFnCreateNewUser) = True Then
        cmdNewUser.Enabled = True
        mbUserHasFnCreateNewUser = True
    Else
        cmdNewUser.Enabled = False
        mbUserHasFnCreateNewUser = False
    End If


    Call EnableApply
'    If goUser.CheckPermission(gsFnChangeAccessRights) = True Then
'        cmdApply.Enabled = True
'    Else
'        cmdApply.Enabled = False
'    End If
    
    If goUser.CheckPermission(gsFnDisableUser) = True Then
        chkEnabled.Enabled = True
    Else
        chkEnabled.Enabled = False
    End If
    
    txtUserName.Enabled = False
    cmdEditUserName.Enabled = False
    'WillC 23/2/2000
    cmdApply.Enabled = False
    
' Turn on key preview for form, so that F1 (Help) can be trapped by form
    Me.KeyPreview = True

'Place code here for the imagelist
    lvwUserDatabase.View = lvwReport  ' Set View property to Report
    lvwUserDatabase.SmallIcons = Img1 ' The icons are held in the Img1 imagelist

' Get the latest data from the database
    RefreshUsers
    RefreshDatabases
    
    
' Add ColumnHeaders with appropriate widths to lvwUserDatabase
     Set clmX = lvwUserDatabase.ColumnHeaders.Add(, , "User Databases", lvwUserDatabase.Width - 110)
    
        For n = 1 To lvwUserDatabase.ListItems.Count
              lvwUserDatabase.ListItems.Item(n).SmallIcon = "BoxNoTick"
        Next n
        
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
Private Sub EnableApply()
'----------------------------------------------------------------------------------------------
'WillC 1/8/00 Check user rights before allowing the Apply button to be enabled.
'----------------------------------------------------------------------------------------------
    
    If goUser.CheckPermission(gsFnChangeAccessRights) = True Then
        cmdApply.Enabled = True
    Else
        cmdApply.Enabled = False
    End If

End Sub
'----------------------------------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------------------------
' Enable the help if F1 is pushed
'----------------------------------------------------------------------------------------------

    If KeyCode = vbKeyF1 Then               ' Show user guide
        'ShowDocument Me.hWnd, gsMACROUserGuidePath
        
        'REM 07/12/01 - New Call to MACRO Help
        Call MACROHelp(Me.hWnd, App.Title)
        
    End If

End Sub

'----------------------------------------------------------------------------------------------
Private Sub cboUsers_Click()
'----------------------------------------------------------------------------------------------
' When you click on a user code in the combo box insert the relevant username in the textbox
'----------------------------------------------------------------------------------------------
Dim sName As String
Dim sUserCode As String
Dim rsUserProfile As ADODB.Recordset
Dim rsEnabled As ADODB.Recordset
Dim n As Integer

    On Error GoTo ErrHandler

    If cboUsers.Text <> "" Then
      sName = suCodeCol.Item(cboUsers.Text)
      txtUserName = sName
    End If

    lvwUserDatabase.Enabled = True
    'changed Mo Morris 22/12/99
    If mbUserHasFnCreateNewUser Then
        cmdEditUserName.Enabled = True
    End If
    
    For n = 1 To lvwUserDatabase.ListItems.Count
      lvwUserDatabase.ListItems.Item(n).SmallIcon = "BoxNoTick"
    Next n

    'Get the Users Details  and choose the relevant items to show the users
    'profile write the Site and Trial details to labels on the form

    sUserCode = cboUsers.Text
    
    
  ' WillC 13/3/00 SR3209 Disable the checkbox if the combo value is the same as the User logged in
  ' so someone cant disable themselves
     
    If LCase(Trim(sUserCode)) = LCase(Trim(goUser.UserName)) Then
        chkEnabled.Enabled = False
        EnableApply
       ' cmdApply.Enabled = False
    Else
        chkEnabled.Enabled = True
        cmdApply.Enabled = False
    End If

    
    Set rsEnabled = gdsUser(sUserCode)
    'WillC 13/3/00 this is a programmatic activated click event not a user one
        mbProgrammaticCheckbox = True
        If rsEnabled!Enabled = 1 Then
            chkEnabled.Value = 1
        Else
            chkEnabled.Value = 0
        End If
   Set rsEnabled = Nothing
        
    
    Set rsUserProfile = UserProfile(sUserCode)
 
    If rsUserProfile.RecordCount = 0 Then
        Exit Sub
    Else
        
     For n = 1 To lvwUserDatabase.ListItems.Count
        lvwUserDatabase.ListItems.Item(n).SmallIcon = "BoxNoTick"
     Next n
        While Not rsUserProfile.EOF = True
             For n = 1 To lvwUserDatabase.ListItems.Count
                'Mo Morris 20/9/01 Db Audit (DatabaseDescription to DatabaseCode)
                If lvwUserDatabase.ListItems.Item(n) = rsUserProfile!DataBaseCode Then
                     lvwUserDatabase.ListItems.Item(n).SmallIcon = "BoxTick"
                End If
            Next n
          rsUserProfile.MoveNext
        Wend
    End If
    
    
    Set rsUserProfile = Nothing
    
    'WillC 13/3/00  Clear the mbProgrammaticCheckbox of any value thats been set.
    mbProgrammaticCheckbox = False
    
Exit Sub
ErrHandler:
    Select Case Err.Number
        Case 3021
            On Error GoTo 0
        Case Else
            Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cboUsers_Click")
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

'----------------------------------------------------------------------------------------------
Private Sub cboUsers_GotFocus()
'----------------------------------------------------------------------------------------------
' Refresh here so that on the insert of a new user the new user can
' be seen immediately
'----------------------------------------------------------------------------------------------
    RefreshUsers

End Sub

'----------------------------------------------------------------------------------------------
Private Sub RefreshUsers()
'----------------------------------------------------------------------------------------------
' Add the list of Users to the User combobox
'----------------------------------------------------------------------------------------------
 Dim rsUsers As ADODB.Recordset
 
    On Error GoTo ErrHandler
    
    cboUsers.Clear
    Set suCodeCol = New Collection
  
   'following line added Mo Morris 7/2/00
   Set rsUsers = New ADODB.Recordset
   Set rsUsers = gdsUserList
   With rsUsers
        Do Until .EOF = True
            ' once the collection is instantiated add the members to it
            ' collection .ADD fields(1) are the names from the recordset and
            ' fields(0) is the usercode which is used as a key to write the data
            ' to the relevant textbox .
            cboUsers.AddItem .Fields(0).Value
            suCodeCol.Add .Fields(1).Value, .Fields(0).Value
                 .MoveNext
        Loop
   End With
   'changed Mo Morris 4/2/00
   rsUsers.Close
   Set rsUsers = Nothing
   
    If suCodeCol.Count > 0 Then
        'TA 12/12/2000: if there is at least one user then select the first
        cboUsers.ListIndex = 0
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "RefreshUsers")
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
Private Sub RefreshDatabases()
'----------------------------------------------------------------------------------------------
' Add the databases to their listview
'----------------------------------------------------------------------------------------------

 'Create a variable to add ListItem objects and receive the list of databases.
Dim itmX As MSComctlLib.ListItem
Dim rsDatabases As ADODB.Recordset

    On Error GoTo ErrHandler

' Get the list of Databases
    Set rsDatabases = gdsDatabaseList
    
    ' While the record is not the last record, add a ListItem object.
    With rsDatabases
        Do Until .EOF = True
          Set itmX = lvwUserDatabase.ListItems.Add(, , .Fields(0).Value)
                itmX.SmallIcon = "BoxNoTick"
            'this places the databasenames in the listview
                 .MoveNext   ' Move to next record.
        Loop
    End With
    
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
Private Sub lvwUserDatabase_Click()
'----------------------------------------------------------------------------------------------
' Allows toggling between the Tick and the NoTick when a user clicks on a database
'----------------------------------------------------------------------------------------------
On Error GoTo ErrHandler
Dim sUserCode As String
    
    'WillC 23/2/2000
    Call EnableApply
    'cmdApply.Enabled = True
    
    sUserCode = cboUsers.Text
        
    'Changed Mo Morris 22/12/99
    If sUserCode <> "" And mbUserHasFnCreateNewUser Then
        cmdEditUserName.Enabled = True
    End If
    
    'Toggle between the Tick and NoTick if ticked add the database to the user
    'if not remove the database from the User
    If lvwUserDatabase.SelectedItem.SmallIcon = "BoxNoTick" Then
            ' AddDatabaseToUser sUserCode, sDatabaseDescription
        lvwUserDatabase.SelectedItem.SmallIcon = "BoxTick"
    Else
       ' RemoveDatabaseFromUser sUserCode, sDatabaseDescription
        lvwUserDatabase.SelectedItem.SmallIcon = "BoxNoTick"
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "lvwUserDatabase_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

