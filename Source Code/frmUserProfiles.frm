VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmUserProfiles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Roles"
   ClientHeight    =   4035
   ClientLeft      =   3585
   ClientTop       =   3255
   ClientWidth     =   5865
   Icon            =   "frmUserProfiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   5415
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete Role"
         Height          =   495
         Left            =   2400
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Close"
         Height          =   495
         Left            =   3960
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin MSFlexGridLib.MSFlexGrid flexUserProfiles 
         Bindings        =   "frmUserProfiles.frx":0442
         Height          =   2055
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   3625
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmUserProfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1999. All Rights Reserved
'   File:       frmUserMaintenance.frm
'   Author:     Will Casey
'   Purpose:    Shows the Users Profiles when they have more than one Role
'--------------------------------------------------------------------------------
'   Revisions:
'   PN  20/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   WillC 18/10/99  Changed the type of grid to allow the deletion of UserRoles
'   WillC 11/10/99  Added the error handlers
'   WillC 4/1/2000  Disabled the delete button until the user chooses a role to be deleted
'   TA 08/05/2000   removed subclassing
'--------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------
Private Sub cmdDelete_Click()
'---------------------------------------------------------------------
'Get all the fields necessary for the entire key needed for the deletion
'---------------------------------------------------------------------
Dim n  As Integer
Dim sMsg As String
Dim sUserCode As String
Dim sRoleCode As String
Dim sDatabaseDescription As String
'Dim sTrialSite As String
'Dim sTrialName As String

    On Error GoTo ErrHandler
          
    sMsg = "Are you sure you want to delete this user role?"
    
    Select Case MsgBox(sMsg, vbQuestion + vbYesNo, gsDIALOG_TITLE)
        Case vbYes
                n = flexUserProfiles.Row
                    flexUserProfiles.Col = 1
                    flexUserProfiles.Row = n
                    sUserCode = flexUserProfiles.Text
                    flexUserProfiles.Text = ""
                    flexUserProfiles.Col = 2
                    flexUserProfiles.Row = n
                    sRoleCode = flexUserProfiles.Text
                   ' flexUserProfiles.Col = 3
                   ' flexUserProfiles.Row = n
                   ' sTrialSite = flexUserProfiles.Text
                   ' flexUserProfiles.Col = 4
                   ' flexUserProfiles.Row = n
                   ' sTrialName = flexUserProfiles.Text
                    flexUserProfiles.Col = 3
                    flexUserProfiles.Row = n
                    sDatabaseDescription = flexUserProfiles.Text
                 
                flexUserProfiles.RemoveItem (n)
                Call DeleteUserRole(sUserCode, sRoleCode, sDatabaseDescription)
               'Call DeleteUserRole(sUserCode, sRoleCode, sTrialSite, sTrialName, sDatabaseDescription)
        Case vbNo
                Exit Sub
    End Select
    
    cmdDelete.Enabled = False
    
Exit Sub
ErrHandler:
    Select Case Err.Number
       '  This error is ok as we wish to allow the user to delete the last userprofile
       '  in the grid and not encounter any problems.
        Case 30015
             Call DeleteUserRole(sUserCode, sRoleCode, sDatabaseDescription)
            flexUserProfiles.Clear
           Unload Me
        Case Else
              Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdDelete_Click")
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

'------------------------------------------------------------------------------'
Private Sub flexUserProfiles_Click()
'------------------------------------------------------------------------------'
' Enable the delete button once a role has been selected
'------------------------------------------------------------------------------'

    cmdDelete.Enabled = True

End Sub

'--------------------------------------------------------------------------------
Private Sub cmdCancel_Click()
'--------------------------------------------------------------------------------
 ' Unload the form
'--------------------------------------------------------------------------------
    
    Unload Me
    
End Sub

'--------------------------------------------------------------------------------
Private Sub Form_Load()
'--------------------------------------------------------------------------------
' Load the grid with all the Roles a user may have set up for them based on their
' Usercode as chosen from the combobox on frmUserMaintenance
'--------------------------------------------------------------------------------
Dim sUserCode As String
 
    On Error GoTo ErrHandler

    If goUser.CheckPermission(gsFnMaintRole) = False Then
        cmdDelete.Enabled = False
    End If


    With flexUserProfiles
        .Top = 200
        .Left = 200
        .Cols = 4
        .MergeCol(0) = True
        .Row = 0
        .ColWidth(0) = 200
        .Col = 1
        .Text = "User code"
        .ColWidth(1) = 1000
        .Col = 2
        .Text = "Role code"
        .ColWidth(2) = 1000
        .Col = 3
        '.Text = "Trial site"
        '.ColWidth(3) = 1000
       ' .Col = 4
       ' .Text = "Trial name"
       ' .ColWidth(4) = 1000
        .Col = 3
        .Text = "Database description"
        .ColWidth(3) = 1800
        .SelectionMode = flexSelectionByRow
    End With
    
    
        sUserCode = frmUserRoles.cboUsers.Text
        GetUserRoles (sUserCode)
    
    'Disable the delete button until a role has been selected
    cmdDelete.Enabled = False
   
    
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

'--------------------------------------------------------------------------------
Private Function GetUserRoles(sUserCode As String)
'--------------------------------------------------------------------------------
' Get the cols you want to display and alias them
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim nTrialId As Long
Dim rsUserProfile As ADODB.Recordset
Dim rsTrialName As ADODB.Recordset
Dim mnRow As Integer
On Error GoTo ErrHandler

    flexUserProfiles.Visible = False
 

'    sSQL = "SELECT Usercode ,RoleCode ," _
'    & "TrialSite ,ClinicalTrialId, DatabaseDescription,AllTrials " _
'    & " FROM UserRole WHERE UserCode = '" & sUserCode & "'" _
'    & " AND AllTrials <= " & 1

    'Mo Morris 20/9/01 Db Audit (DatabaseDescription to DatabaseCode, UserCode to UserName)
    sSQL = "SELECT UserName ,RoleCode ," _
    & " DatabaseCode" _
    & " FROM UserRole WHERE UserName = '" & sUserCode & "'"
    
    
    Set rsUserProfile = New ADODB.Recordset
    rsUserProfile.CursorLocation = adUseClient
    rsUserProfile.Open sSQL, SecurityADODBConnection, adOpenKeyset, , adCmdText
    If rsUserProfile.EOF = True Then
        cmdDelete.Enabled = False
    End If
    
    mnRow = 1
    Do While Not rsUserProfile.EOF
        flexUserProfiles.Rows = mnRow + 1
        flexUserProfiles.Row = mnRow
        flexUserProfiles.Col = 0
        flexUserProfiles.Text = ""
        flexUserProfiles.CellBackColor = flexUserProfiles.BackColor
        flexUserProfiles.Col = 1
        flexUserProfiles.Text = rsUserProfile!UserName
        flexUserProfiles.Col = 2
        flexUserProfiles.Text = rsUserProfile!RoleCode
    '    flexUserProfiles.Col = 3
    '    flexUserProfiles.Text = rsUserProfile!TrialSite
    '    flexUserProfiles.Col = 4
    '    If rsUserProfile!AllTrials = 1 Then
    '         flexUserProfiles.Text = "All Trials"
    '    Else
    '        flexUserProfiles.Text = rsTrialName!ClinicalTrialName
    '    End If
        flexUserProfiles.Col = 3
        flexUserProfiles.Text = rsUserProfile!DataBaseCode
        mnRow = mnRow + 1
        rsUserProfile.MoveNext
    Loop

    flexUserProfiles.Visible = True
    Set rsUserProfile = Nothing
    Set rsTrialName = Nothing
    
Exit Function
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "GetUserRoles")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Function

'--------------------------------------------------------------------------------

Private Sub DeleteUserRole(sUserCode As String, sRoleCode As String, sDatabaseDescription As String) 'sTrialSite As String, _
                                    sTrialName As String, sDatabaseDescription As String)
'--------------------------------------------------------------------------------
' Delete the UserRole from the security database, first get the trialid from the
' macro Database to make sure we are deleting on the entire primary key.
' trap for a trial being "all Trials" if so set the trialId to zero.
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim nTrialId As Long
Dim rsTrialID As ADODB.Recordset
Dim n As Integer
On Error GoTo ErrHandler
                                                                 
    'Changed by Mo Morris 18/1/00. '*' removed from Delete SQL statement
    'Mo Morris 20/9/01 Db Audit (DatabaseDescription to DatabaseCode, UserCode to UserName)
    sSQL = " DELETE FROM UserRole " _
      & " WHERE UserName = '" & sUserCode & "'" _
      & " AND RoleCode = '" & sRoleCode & "'" _
      & " AND DatabaseCode = '" & sDatabaseDescription & "'"

'      & " AND TrialSite = '" & sTrialSite & "'" _
'      & " AND ClinicalTrialId = " & nTrialId _
'      & " AND DatabaseDescription = '" & sDatabaseDescription & "'"
    SecurityADODBConnection.Execute sSQL, , adCmdText
    sUserCode = frmUserRoles.cboUsers.Text
    GetUserRoles (sUserCode)
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "DeleteUserRole")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

