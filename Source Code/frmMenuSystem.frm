VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MACRO System Management"
   ClientHeight    =   6525
   ClientLeft      =   1890
   ClientTop       =   1920
   ClientWidth     =   8340
   ClipControls    =   0   'False
   Icon            =   "frmMenuSystem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   8340
   Begin VB.CommandButton cmdResetPassword 
      Caption         =   "Reset &Other Password"
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   1800
      Width           =   2895
   End
   Begin VB.CommandButton cmdRestoreDB 
      Caption         =   "Res&tore Site Database"
      Height          =   615
      Left            =   3720
      TabIndex        =   11
      Top             =   4320
      Width           =   2895
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   615
      Left            =   600
      TabIndex        =   12
      Top             =   5160
      Width           =   2895
   End
   Begin VB.CommandButton cmdDataCommunications 
      Caption         =   "Data &Communication"
      Height          =   615
      Left            =   3720
      TabIndex        =   9
      Top             =   3480
      Width           =   2895
   End
   Begin VB.CommandButton cmdLockManagement 
      Caption         =   "Database &Lock Administration"
      Height          =   615
      Left            =   600
      TabIndex        =   10
      Top             =   4320
      Width           =   2895
   End
   Begin VB.CommandButton CmdRoleMgmt 
      Caption         =   "&Role Management..."
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton cmdUsers 
      Caption         =   "&Users..."
      Height          =   615
      Left            =   3720
      TabIndex        =   3
      Top             =   960
      Width           =   2895
   End
   Begin VB.CommandButton cmdUserRoles 
      Caption         =   "Us&er Roles..."
      Height          =   615
      Left            =   3720
      TabIndex        =   5
      Top             =   1800
      Width           =   2895
   End
   Begin VB.CommandButton cmdPwdSettings 
      Caption         =   "Password &Settings"
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   2895
   End
   Begin VB.CommandButton cmdChangePwd 
      Caption         =   "Change &My Password"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton cmdNewDatabase 
      Caption         =   "  &Databases..."
      Height          =   615
      Left            =   3720
      TabIndex        =   7
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Timer tmrSystemIdleTimeout 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdProperties 
      Caption         =   "&Properties"
      Height          =   615
      Left            =   600
      TabIndex        =   6
      Top             =   2640
      Width           =   2895
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   3720
      TabIndex        =   13
      Top             =   5160
      Width           =   2895
   End
   Begin VB.CommandButton cmdLog 
      Caption         =   "System &Log"
      Height          =   615
      Left            =   600
      TabIndex        =   8
      Top             =   3480
      Width           =   2895
   End
   Begin VB.CommandButton cmdSecurity 
      Caption         =   "Security &control"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7080
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar sbrMenu 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   15
      Top             =   6090
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "UserKey"
            Object.ToolTipText     =   "Name of current user"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "RoleKey"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Bevel           =   0
            Object.Width           =   1766
            MinWidth        =   1766
            TextSave        =   "14/10/2002"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "11:11"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   5927
            Key             =   "Key"
            Object.ToolTipText     =   "Key for data status symbols"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   2
            Object.Visible         =   0   'False
            Text            =   "Display symbol key"
            TextSave        =   "Display symbol key"
            Key             =   "DisplayKey"
            Object.ToolTipText     =   "Click to display the key"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblUsername 
      Caption         =   "Label1"
      Height          =   375
      Left            =   4560
      TabIndex        =   16
      Top             =   7440
      Width           =   2295
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuArezzoSettings 
         Caption         =   "A&REZZO Memory Settings"
      End
      Begin VB.Menu mnuFAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuFExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       frmMenuSystem.frm
'   Author:     Andrew Newbigging, November 1997
'   Purpose:    Main menu used in Macro System Management.
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
'   Revisions:
'   1    Andrew Newbigging    21/11/97
'   2    Andrew Newbigging    27/11/97
'   3    Andrew Newbigging    27/01/98
'
'   NCJ 16 Sept 99
'       Added InitialiseMe routine
'   PN  20/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   PN  23/09/99    Added tmrSystemShutdown timer control to manage proper shutdown
'                   of all forms in system, including modally displayed forms
'   NCJ 30/9/99     Added declaration for mbSystemLocked
'   Mo Morris   15/12/99    Property Get Mode (=gsSYSTEM_MANAGER_MODE) added
'   NCJ 13/1/00     Set gsMACROUserGuidePath in InitialiseMe
'   TA 08/05/2000   removed subclassing
'   TA 01/08/2000 SR3652   form caption now title case
'   WillC 1/8/00 SR3728 Added frmResetPwd to show the form to change other passwords
'   NCJ 27/9/00 - Fixed bug in enabling "Change Password" button
'   TA 04/12/00 - Standardised case of button captions
'   TA 13/12/00 - Changed tab ordering
'   DPH 22/10/2001 - Lock Admin Added
'   DPH 26/10/2001 - Password Change Fix
'---------------------------------------------------------------------
'
'--------------------------------------------------------------------------------

Option Explicit
Option Base 0
Option Compare Binary


Private mbSystemLocked As Boolean

'---------------------------------------------------------------------
Private Sub cmdChangePwd_Click()
'---------------------------------------------------------------------
' bring up the change password dialogue
'---------------------------------------------------------------------
' REVISIONS
' DPH 26/10/2001 - Place new password into user class
'---------------------------------------------------------------------
Dim sMessage As String
Dim bChangePassword As Boolean

    'TA 7/6/2000 SR359:  updated so display function is called)
    'REM 10/10/02 - Changed to new form
    If frmPasswordChange.Display(goUser, , " ") Then
        Call goUser.gLog(goUser.UserName, "ChangePassword", "System admin changed user password")
    End If

    ' DPH 26/10/2001 - storing any password change in class member
'    If sNewPassword <> "" Then
'        gUser.Password = sNewPassword
'    End If

End Sub
'---------------------------------------------------------------------
Private Sub cmdDataCommunications_Click()
'---------------------------------------------------------------------
' show the data communication form
'---------------------------------------------------------------------
    
    'WillC 22/6/00 added vbmodal due to message box in form
    frmDataCommunication.Show vbModal

End Sub

'---------------------------------------------------------------------
Private Sub cmdAbout_Click()
'---------------------------------------------------------------------
' Show the about screen
'---------------------------------------------------------------------
    On Error GoTo ErrHandler
        
        frmAbout.Show
        
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdAbout_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub cmdHelp_Click()
'---------------------------------------------------------------------
' show the helpfile
' NCJ 20 Jan 00 - Removed "No help file" comment
'---------------------------------------------------------------------

    ' gsMACROUserGuidePath is set up in InitialiseMe
    'Call ShowDocument(Me.hWnd, gsMACROUserGuidePath)

    'REM 07/12/01 - New Call to MACRO Help
    Call MACROHelp(Me.hWnd, App.Title)

End Sub

'---------------------------------------------------------------------
Private Sub cmdLockManagement_Click()
'---------------------------------------------------------------------
' DPH 22/10/2001 - Added Lock Management
'---------------------------------------------------------------------
    On Error GoTo ErrHandler

     Call frmLocksAdmin.Display(goUser)
     
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdLockManagement_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
End Sub

'---------------------------------------------------------------------
Private Sub cmdNewDatabase_Click()
'---------------------------------------------------------------------
' show the form to allow new databases or editing of existing
'---------------------------------------------------------------------
    On Error GoTo ErrHandler
    
   frmDatabases.Show vbModal
   'ash 15/09/2002 test
   'frmDatabaseInfo.Show
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdNewDatabase_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
   
End Sub

'---------------------------------------------------------------------
Private Sub cmdPwdSettings_Click()
'---------------------------------------------------------------------
'bring up the form to change password settings
'---------------------------------------------------------------------
Dim sMessage As String

    Call frmPasswordPolicy.Display
 
End Sub

'---------------------------------------------------------------------
Private Sub cmdRestoreDB_Click()
'---------------------------------------------------------------------
' Show the Restore Site Database form
'---------------------------------------------------------------------
    frmRestoreSiteDatabase.Show vbModal


End Sub

'---------------------------------------------------------------------
Private Sub CmdRoleMgmt_Click()
'---------------------------------------------------------------------
' Show the role mgmt form
'---------------------------------------------------------------------

    frmRoleMgmt.Show vbModal
    'ash 18/9/2002 test
    'frmRoleManagement.Show vbModal
 
End Sub

'---------------------------------------------------------------------
Private Sub cmdUserRoles_Click()
'---------------------------------------------------------------------
'bring up the form to see user settings
'---------------------------------------------------------------------

    'frmUserRoles.Show vbModal
    
    'ash 17/9/2002 test
    'frmNWUserRole.Show vbModal
    
    frmNWUserRole.Display
 
End Sub

'---------------------------------------------------------------------
Private Sub cmdUsers_Click()
'---------------------------------------------------------------------
'bring up the form to see user settings
'---------------------------------------------------------------------

    frmUsers.Show vbModal
 
End Sub

'---------------------------------------------------------------------
Private Sub cmdResetPassword_Click()
'---------------------------------------------------------------------
'bring up the form to change other passwords
'---------------------------------------------------------------------
'WillC 1/8/00 SR3728
    frmResetPwd.Display
    
End Sub

'------------------------------------------------------------------------'
Private Sub Form_Load()
'------------------------------------------------------------------------'

'------------------------------------------------------------------------'
    FormCentre Me
    
End Sub

'------------------------------------------------------------------------------'
Private Sub Form_Unload(Cancel As Integer)
'------------------------------------------------------------------------'
    

    Call ExitMACRO
    'MACROEnd added by Mo Morris 7/2/00
    'And then taken out again, because it caused severe crashes
    'Call MACROEnd
 
End Sub

'--------------------------------------
Public Sub InitialiseMe()
'--------------------------------------
' Perform initialisations specific to System Management
' Called from Main in MainMacroModule
' NCJ 16 Sept 1999
'--------------------------------------

    ' NCJ 13/1/00 - Set System Management user guide
    gsMACROUserGuidePath = gsMACROUserGuidePath & "SM\Contents.htm"
    
End Sub

'--------------------------------------
Private Sub cmdExit_Click()
'--------------------------------------
    'logg user offf
    goUser.LogOff
    Unload Me
 
End Sub

'--------------------------------------
Private Sub cmdLog_Click()
'--------------------------------------
    
   frmLogDetails.Show vbModal
 
End Sub

'--------------------------------------
Private Sub cmdProperties_Click()
'--------------------------------------
' PN 17/09/99 - created
' display the system properties window
'--------------------------------------

    frmSystemProperties.Show vbModal
    
End Sub

'--------------------------------------
Private Sub cmdSecurity_Click()
'--------------------------------------

    'frmUserMaintenance.Show vbModal

End Sub

'---------------------------------------------------------------------
Private Sub mnuArezzoSettings_Click()
'---------------------------------------------------------------------
'loads AREZZO settings form
'---------------------------------------------------------------------
    
    frmArezzoSettings.Show vbModal
End Sub

'---------------------------------------------------------------------
Private Sub mnuFAbout_Click()
'---------------------------------------------------------------------
' DPH 22/10/2001 - About button moved to Menu item to make room for
' Lock Management Button
'---------------------------------------------------------------------

    cmdAbout_Click
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuFExit_Click()
'---------------------------------------------------------------------
' DPH 22/10/2001 - Exit added to menu as About button placed there
'---------------------------------------------------------------------

    cmdExit_Click
    
End Sub



'---------------------------------------------------------------------
Private Sub tmrSystemIdleTimeout_Timer()
'---------------------------------------------------------------------
' nb This timer event should never occur unless DevMode = 0
' when the timer goes off it must be time to lock the system
'  it prompts the user to enter the password or wxit MACRO
' the system is then either closed in a controlled way or resets the timer
' NCJ 17/3/00 - Tidied up and simplified (SR 3015)
'TA 27/04/2000: new timeout handling
'---------------------------------------------------------------------
    'new timeout handling
    glSystemIdleTimeoutCount = glSystemIdleTimeoutCount + 1
    If glSystemIdleTimeout = glSystemIdleTimeoutCount Then
        ' set the couter to 0 and disable the timer until the user logs in
        glSystemIdleTimeoutCount = 0
        tmrSystemIdleTimeout.Enabled = False
        If frmTimeOutSplash.Display Then
            'password correctly entered
            tmrSystemIdleTimeout.Enabled = True
        Else
            'exit MACRO chosen
            ' unload all forms and exit
            Call UnloadAllForms
        End If
    End If
    
End Sub

'---------------------------------------------------------------------
 Public Sub CheckUserRights()
'---------------------------------------------------------------------
' Check to see if the user has certain rights
' DPH 22/10/2001 - Lock Admin Added
'---------------------------------------------------------------------
    
    CmdRoleMgmt.Enabled = goUser.CheckPermission(gsFnMaintRole)
    
    ' NCJ 27/9/00 - Corrected to cmdChangePwd
    cmdChangePwd.Enabled = goUser.CheckPermission(gsFnChangePassword)
    
    'added by Mo Morris 21/12/99
    If goUser.CheckPermission(gsFnChangeAccessRights) = True _
        And goUser.CheckPermission(gsFnAssignUserToTrial) = True Then
            cmdUserRoles.Enabled = True
    Else
            cmdUserRoles.Enabled = False
    End If
    
    'WillC SR3728 1/8/00
    cmdResetPassword.Enabled = goUser.CheckPermission(gsFnResetPassword)

    cmdPwdSettings.Enabled = goUser.CheckPermission(gsFnChangeSystemProperties)
    
    cmdProperties.Enabled = goUser.CheckPermission(gsFnChangeSystemProperties)
    
    cmdLog.Enabled = goUser.CheckPermission(gsFnViewSystemLog)
    
    cmdDataCommunications.Enabled = goUser.CheckPermission(gsFnViewSiteServerCommunication)
    
    cmdRestoreDB.Enabled = goUser.CheckPermission(gsFnRestoreDatabase)
    
    
    If goUser.CheckPermission(gsFnChangeAccessRights) Or goUser.CheckPermission(gsFnCreateNewUser) Or goUser.CheckPermission(gsFnDisableUser) Then
        cmdUsers.Enabled = True
    Else
        cmdUsers.Enabled = False
    End If

    If goUser.CheckPermission(gsFnCreateDB) Or goUser.CheckPermission(gsFnRegisterDB) Then
        cmdNewDatabase.Enabled = True
    Else
        cmdNewDatabase.Enabled = False
    End If

    ' DPH 22/10/2001 - Lock Admin Permissions
    If goUser.CheckPermission(gsFnRemoveAllLocks) Or goUser.CheckPermission(gsFnRemoveOwnLocks) Then
        cmdLockManagement.Enabled = True
    Else
        cmdLockManagement.Enabled = False
    End If
 End Sub
 
 '---------------------------------------------------------------------
Public Property Get Mode() As String
'---------------------------------------------------------------------

    Mode = gsSYSTEM_MANAGER_MODE

End Property
