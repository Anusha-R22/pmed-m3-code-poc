VERSION 5.00
Begin VB.Form frmUserRoles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Roles"
   ClientHeight    =   4860
   ClientLeft      =   3795
   ClientTop       =   3465
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboDatabases 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1200
      Width           =   3375
   End
   Begin VB.OptionButton optAllTrialsAllSites 
      Caption         =   "Apply role to all Trials at all Sites"
      Height          =   330
      Left            =   5160
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton optAllSitesATrial 
      Caption         =   "Apply role to all Sites for a Trial"
      Height          =   330
      Left            =   7680
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.OptionButton optAllTrialsASite 
      Caption         =   "Apply role to all Trials at a Site"
      Height          =   330
      Left            =   7080
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.OptionButton optASiteATrial 
      Caption         =   "Apply role to a single Site and Trial"
      Height          =   330
      Left            =   6120
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      Width           =   6975
      Begin VB.CommandButton cmdDeleteRoles 
         Caption         =   "&Delete roles..."
         Height          =   495
         Left            =   1680
         TabIndex        =   21
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Close"
         Height          =   495
         Left            =   5040
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdUserProfile 
         Caption         =   "&View user roles..."
         Height          =   495
         Left            =   3360
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.ComboBox cboRoles 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   6975
      Begin VB.TextBox txtDatabase 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   20
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtUserRole 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Top             =   840
         Width           =   3375
      End
      Begin VB.CommandButton cmdApplyRole 
         Caption         =   "&Apply Role..."
         Height          =   495
         Left            =   5040
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "User role:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "User database:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.ComboBox cboUsers 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin VB.TextBox txtUserName 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "User roles:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "User database:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Full real name:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label lblUserCode 
      Caption         =   "User name:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmUserRoles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       frmUserRoles.frm
'   Author:     Will Casey 22/10/99
'   Purpose:    Maintain users and the databases which a user can access.
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
'   Revisions:
' WillC     11/10/99 Added error handler
' WillC     21/2/00  Changed cmdApplyRole a User now has one role per database
'   WillC 7/8/00 SR2956   Changed labels
'--------------------------------------------------------------------------------

Option Explicit
Option Compare Binary
Option Base 0
' this collection holds the usernames and usercodes the other FunctionCodes
Private suCodeCol As Collection
Private mcolUserFunction As New Collection

'---------------------------------------------------------------------
Private Sub cmdEditUserName_Click()
'---------------------------------------------------------------------
' show the form
'---------------------------------------------------------------------
    frmEditUserName.Show vbModal

End Sub

'---------------------------------------------------------------------
Private Sub cboDatabases_Click()
'---------------------------------------------------------------------
' Fill the text box with the combo box item
'---------------------------------------------------------------------

    txtDatabase.Text = cboDatabases.Text
       
End Sub

'---------------------------------------------------------------------
Private Sub cboRoles_Click()
'---------------------------------------------------------------------
' Fill the text box with the combo box item
'---------------------------------------------------------------------

    txtUserRole.Text = cboRoles.Text

End Sub
'---------------------------------------------------------------------
Private Sub cboUsers_Change()
'---------------------------------------------------------------------
' if theres no selection in the combo clear the text box
'---------------------------------------------------------------------
        
        If cboUsers.Text = "" Then
            txtUserName = ""
        End If
End Sub

'----------------------------------------------------------------------------------------------
Private Sub cboUsers_Click()
'----------------------------------------------------------------------------------------------
' When you click on a user code in the combo box insert the relevant username in the textbox
'----------------------------------------------------------------------------------------------
Dim sName As String
    
      'having declared the cboUsersText(userCode) as the key below in RefreshUsers
      'use it as the key to retrieve the corresponding  username
    If cboUsers.Text = "" Then
        txtUserName = ""
    Else
      sName = suCodeCol.Item(cboUsers.Text)
      txtUserName = sName
    End If
  
    cmdUserProfile.Enabled = True
    cmdDeleteRoles.Enabled = True

    'following lines added by Mo Morris 21/12/99
    If cboUsers.Text <> "" _
    And (txtDatabase.Text <> "") _
    And (txtUserRole.Text <> "") Then
        cmdApplyRole.Enabled = True
        cmdDeleteRoles.Enabled = True
    Else
        cmdApplyRole.Enabled = False
    End If
   
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cboUsers_Click")
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
Private Sub cmdDeleteRoles_Click()
'----------------------------------------------------------------------------------------------
' Show the form UserProfiles
'----------------------------------------------------------------------------------------------

    cmdUserProfile.Value = True
    
End Sub

'----------------------------------------------------------------------------------------------
Private Sub cmdUserProfile_Click()
'----------------------------------------------------------------------------------------------
' If a user has more than one role then show the UserProfile form , however
' if a user has only one role show the role on the main form..
'----------------------------------------------------------------------------------------------
    
Dim sUserCode As String
Dim rsUserProfile As ADODB.Recordset
Dim n As Integer

    On Error GoTo ErrHandler

    'Get the Users Details  and choose the relevant items to show the users
    'profile write the Site and Trial details to labels on the form

    sUserCode = cboUsers.Text
    
    Set rsUserProfile = UserProfile(sUserCode)
 
 ' Move to the last record so you can do a count, if its more than one record show
 ' them on the form UserProfiles if its just one  then show the profile on this form
    If Not rsUserProfile.EOF Then
       rsUserProfile.MoveLast
    End If
 
        frmUserProfiles.Show vbModal
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdUserProfile_Click")
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
Private Sub cboUsers_GotFocus()
'----------------------------------------------------------------------------------------------
' Refresh here so that on the insert of a new user the new user can
' be seen immediately
'----------------------------------------------------------------------------------------------
    RefreshUsers
    
End Sub
'----------------------------------------------------------------------------------------------
Private Sub Form_Load()
'----------------------------------------------------------------------------------------------
'Open the Database, and Load all the Lists with the data, add the columns to the
' listview, Disable any buttons that could cause a problem at this stage.
'----------------------------------------------------------------------------------------------
 
    On Error GoTo ErrHandler
      
    Me.Icon = frmMenu.Icon
    
    txtDatabase.Text = ""
    txtUserName.Text = ""
    txtUserRole.Text = ""
    ' Turn on key preview for form, so that F1 (Help) can be trapped by form
    Me.KeyPreview = True
    
    ' Get the latest data from the database
    RefreshRoles
    RefreshUsers
    RefreshDatabases
    
    'txtDatabase.Text = gUser.DatabaseName
    txtDatabase.Enabled = False
    
    'Disable the Command button stops runtime errors enabled on the click of cboUsers
    cmdUserProfile.Enabled = False
    cmdApplyRole.Enabled = False
    cmdDeleteRoles.Enabled = False
        
    Call FormCentre(Me)
    
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
Private Sub RefreshUsers()
'----------------------------------------------------------------------------------------------
' Add the list of Users to the User combobox
'----------------------------------------------------------------------------------------------
 Dim rsUsers As ADODB.Recordset

   On Error GoTo ErrHandler

   cboUsers.Clear
   Set suCodeCol = New Collection
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
Private Sub RefreshRoles()
'----------------------------------------------------------------------------------------------
 ' Add the roles to their listbox
'----------------------------------------------------------------------------------------------
Dim rsRoles As ADODB.Recordset

    On Error GoTo ErrHandler

    cboRoles.Clear
     
    ' Get the list of roles
    Set rsRoles = gdsRoleList
    With rsRoles
       Do Until .EOF = True
         'this adds the RoleCode field from the table to the list box
         cboRoles.AddItem .Fields(0).Value
                .MoveNext
       Loop
    End With
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "RefreshRoles")
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
'Private Sub RefreshDatabases()
''----------------------------------------------------------------------------------------------
'' Add the databases to their listview
''----------------------------------------------------------------------------------------------
' 'Create a variable to add ListItem objects and receive the list of databases.
'Dim rsDatabases As ADODB.Recordset
'On Error GoTo ErrHandler
'
'' Get the list of Databases
'    Set rsDatabases = gdsDatabaseList
'
'    ' While the record is not the last record, add a ListItem object.
'    With rsDatabases
'        Do Until .EOF = True
'           cboDatabases.AddItem (rsDatabases!DataBaseDescription)
'                 .MoveNext   ' Move to next record.
'        Loop
'    End With
'
'Exit Sub
'ErrHandler:
'   Call MACROFormErrorHandler(Me, Err.Number, Err.Description)
'
'End Sub


'----------------------------------------------------------------------------------------------
Private Sub cmdApplyRole_Click()
'----------------------------------------------------------------------------------------------
' Apply the selection made by the choice of option button optAllTrialsAllSites
' as we wont have a screen for that as there are no choices to be made, ie this role for
' all trials at all sites.
' TA 31/10/2000 - message boxes text and titles tidied up
'----------------------------------------------------------------------------------------------
Dim sUserCode As String
Dim sRoleCode As String
Dim sDatabaseDescription As String
Dim sUsername As String
Dim sMsg As String

    On Error GoTo ErrHandler
     
    sDatabaseDescription = txtDatabase.Text
    sRoleCode = txtUserRole.Text
    sUserCode = cboUsers.Text
    sUsername = txtUserName.Text
    
    ' WillC  21/2/00 SR2853  A User now has one role per database
      If gblnUserRoleOnDatabaseExists(sUserCode, sDatabaseDescription) Then
     'If gblnUserRoleExists(sUserCode, sRoleCode, sDatabaseDescription) Then 'sRoleCode,, 0, sTrialSite) Then
            sMsg = " This user has an existing role on this database." & vbCrLf _
                 & " Do you wish to overwrite the existing role?"
            If DialogQuestion(sMsg) = vbYes Then
                Call gdsUpdateAllTrialsAllSites(sUserCode, sRoleCode, sDatabaseDescription)
            Else
                txtDatabase.Text = ""
                txtUserRole.Text = ""
                '3 lines added by Mo Morris 21/12/99
                cboUsers.ListIndex = -1
                cboDatabases.ListIndex = -1
                cboRoles.ListIndex = -1
                cmdApplyRole.Enabled = False
                Exit Sub
            End If
       Else
            sMsg = "You are about to give " & sUsername & " the role " & sRoleCode & " on the " & sDatabaseDescription & " database." & vbCrLf _
                      & "Do you wish to continue?"
            If DialogQuestion(sMsg) = vbYes Then
                    Call gdsInsertAllTrialsAllSites(sUserCode, sRoleCode, sDatabaseDescription)
            Else
                  ' just unload
                  ' so do nothing here
            End If
     End If
     
    txtDatabase.Text = ""
    txtUserRole.Text = ""
    '3 lines added by Mo Morris 21/12/99
    cboUsers.ListIndex = -1
    cboDatabases.ListIndex = -1
    cboRoles.ListIndex = -1
    
   cmdApplyRole.Enabled = False
   cmdUserProfile.Enabled = False
   cmdDeleteRoles.Enabled = False
   
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdApplyRole_Click")
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

    Unload Me
    
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
Private Sub RefreshDatabases()
'----------------------------------------------------------------------------------------------
' Add the databases to their listview
'----------------------------------------------------------------------------------------------
 'Create a variable to add ListItem objects and receive the list of databases.
Dim rsDatabases As ADODB.Recordset

    On Error GoTo ErrHandler

    ' Get the list of Databases
    Set rsDatabases = gdsDatabaseList
    ' While the record is not the last record, add a ListItem object.
    With rsDatabases
        Do Until .EOF = True
            'Mo Morris 20/9/01 Db Audit (DatabaseDescription to DatabaseCode)
            cboDatabases.AddItem (rsDatabases!DataBaseCode)
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
Private Sub txtDatabase_Change()
'----------------------------------------------------------------------------------------------
're-written by Mo Morris 21/12/99
'----------------------------------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If cboUsers.Text <> "" _
    And (txtDatabase.Text <> "") _
    And (txtUserRole.Text <> "") Then
        cmdApplyRole.Enabled = True
    Else
        cmdApplyRole.Enabled = False
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtDatabase_Change")
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
Private Sub txtUserRole_Change()
'----------------------------------------------------------------------------------------------
're-written by Mo Morris 21/12/99
'----------------------------------------------------------------------------------------------

    On Error GoTo ErrHandler

    If cboUsers.Text <> "" _
    And (txtDatabase.Text <> "") _
    And (txtUserRole.Text <> "") Then
        cmdApplyRole.Enabled = True
    Else
        cmdApplyRole.Enabled = False
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtUserRole_Change")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'Commented out by Mo Morris, Checking Rights is now done in frmMenu.CheckUserRights
''----------------------------------------------------------------------------------------------
'Public Function CheckRightsOK() As Boolean
''----------------------------------------------------------------------------------------------
'' disable buttons where user has no rights
''----------------------------------------------------------------------------------------------
'
'On Error GoTo ErrHandler
'
'    If gUser.CheckFunctionAccess(gsFnChangeAccessRights) = False Then
'        cmdApplyRole.Enabled = False
'    Else
'        cmdApplyRole.Enabled = True
'    End If
'
'    If gUser.CheckFunctionAccess(gsFnAssignUserToTrial) = False Then
'        cmdApplyRole.Enabled = False
'    Else
'        cmdApplyRole.Enabled = True
'    End If
'
'    If gUser.CheckFunctionAccess(gsFnMaintRole) = False Then
'        cmdApplyRole.Enabled = False
'    Else
'        cmdApplyRole.Enabled = True
'    End If
'
'Exit Function
'ErrHandler:
'  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtUserRole_Change")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Unload frmMenu
'   End Select
'End Function
