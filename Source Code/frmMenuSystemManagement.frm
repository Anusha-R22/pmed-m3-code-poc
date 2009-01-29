VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMenu 
   BackColor       =   &H8000000C&
   Caption         =   "MACRO System Management"
   ClientHeight    =   8550
   ClientLeft      =   1695
   ClientTop       =   2610
   ClientWidth     =   11355
   Icon            =   "frmMenuSystemManagement.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   Begin VB.Timer tmrSystemIdleTimeout 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.StatusBar sbrMenu 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   8235
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Key             =   "UserKey"
            Object.ToolTipText     =   "Name of current user"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Text            =   "Role"
            TextSave        =   "Role"
            Key             =   "RoleKey"
            Object.ToolTipText     =   "Name of current users role"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Key             =   "UserDatabase"
            Object.ToolTipText     =   "Name of current database."
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Bevel           =   0
            TextSave        =   "28/02/2008"
            Object.ToolTipText     =   "Date"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Bevel           =   0
            TextSave        =   "14:21"
            Object.ToolTipText     =   "Time"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Enabled         =   0   'False
      Begin VB.Menu mnuUPasswordPolicy 
         Caption         =   "&Password Policy..."
      End
      Begin VB.Menu mnuUSystemPolicy 
         Caption         =   "&System Properties..."
      End
      Begin VB.Menu mnuActiveDirectoryServers 
         Caption         =   "Active &Directory Servers..."
      End
      Begin VB.Menu mnuUSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogOff 
         Caption         =   "&Logout"
      End
      Begin VB.Menu mnuUExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "&Database"
      Enabled         =   0   'False
      Begin VB.Menu mnuDCreateDatabase 
         Caption         =   "&Create Database..."
      End
      Begin VB.Menu mnuDRegisterDatabase 
         Caption         =   "&Register Database..."
      End
      Begin VB.Menu mnuDunRegisterDatabase 
         Caption         =   "&Unregister Database"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuUpgradeDatabase 
         Caption         =   "&Upgrade Database"
      End
      Begin VB.Menu mnuDSetDatabasePassword 
         Caption         =   "Change Database &Password..."
      End
      Begin VB.Menu mnuHTMLFolder 
         Caption         =   "C&hange HTML Folder..."
      End
      Begin VB.Menu mnuDSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDSecurityDatabase 
         Caption         =   "&Security Database..."
      End
      Begin VB.Menu mnuDSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDRestoreSiteDatabase 
         Caption         =   "Restore Site &Database..."
      End
      Begin VB.Menu mnuDSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDLockAdministration 
         Caption         =   "&Lock Administration..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDTimezone 
         Caption         =   "Database &Timezone..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuUser 
      Caption         =   "&User"
      Enabled         =   0   'False
      Begin VB.Menu mnuUNewUser 
         Caption         =   "&New User..."
      End
      Begin VB.Menu mnuUGoToUser 
         Caption         =   "&Go to User"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuUSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUNewUserRole 
         Caption         =   "Ne&w User Role..."
      End
      Begin VB.Menu mnuUDeleteUserRole 
         Caption         =   "&Delete User Role"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuUSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUResetPassword 
         Caption         =   "&Reset My Password..."
      End
   End
   Begin VB.Menu mnuRole 
      Caption         =   "&Role"
      Enabled         =   0   'False
      Begin VB.Menu mnuRNewRole 
         Caption         =   "&New Role..."
      End
      Begin VB.Menu mnuRGoToRole 
         Caption         =   "&Go to Role"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Enabled         =   0   'False
      Begin VB.Menu mnuVRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuVSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVCommunicationLog 
         Caption         =   "&Communication Log..."
      End
      Begin VB.Menu mnuVSystemLog 
         Caption         =   "&System Log..."
      End
      Begin VB.Menu mnuLoginLog 
         Caption         =   "&User Log..."
      End
      Begin VB.Menu mnuVSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVUnusableRoleAssignments 
         Caption         =   "&Display unassigned nodes"
      End
   End
   Begin VB.Menu mnuCurrentDB 
      Caption         =   "&Tasks"
      Enabled         =   0   'False
      Begin VB.Menu mnuSite 
         Caption         =   "&Site"
         Begin VB.Menu mnuSiteAdministration 
            Caption         =   "Site Administration..."
         End
         Begin VB.Menu mnuStudySiteAdministration 
            Caption         =   "Study Site Administration..."
         End
         Begin VB.Menu mnuLaboratorySiteAdministration 
            Caption         =   "Laboratory Site Administration..."
         End
      End
      Begin VB.Menu mnuStudy 
         Caption         =   "S&tudy"
         Begin VB.Menu mnuExportStudyDefinition 
            Caption         =   "&Export..."
         End
         Begin VB.Menu mnuImportStudyDefinition 
            Caption         =   "&Import..."
         End
         Begin VB.Menu mnuStudyStatus 
            Caption         =   "Status..."
         End
         Begin VB.Menu mnuGenerateHTML 
            Caption         =   "Generate HTML..."
         End
      End
      Begin VB.Menu mnuLaboratory 
         Caption         =   "&Laboratory"
         Begin VB.Menu mnuExport 
            Caption         =   "Export..."
         End
         Begin VB.Menu mnuImport 
            Caption         =   "Import..."
         End
         Begin VB.Menu mnuDistribute 
            Caption         =   "Distribute..."
         End
      End
      Begin VB.Menu mnuSubject 
         Caption         =   "S&ubject"
         Begin VB.Menu mnuSuExport 
            Caption         =   "Export..."
         End
         Begin VB.Menu mnuSuImport 
            Caption         =   "Import..."
         End
      End
      Begin VB.Menu mnuVSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDataComm 
         Caption         =   "Data Communication"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Enabled         =   0   'False
      Begin VB.Menu mnuHUserGuide 
         Caption         =   "&User Guide"
      End
      Begin VB.Menu mnuVSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHAbout 
         Caption         =   "&About MACRO"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUpSubItem 
         Caption         =   "PopUpSubITem"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
'   File:       frmMenuSystemManagement.frm
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   Author:     Matthew Martin, October 2002
'   Purpose:    Main menu form for MACRO 3.0 System Management module
'------------------------------------------------------------------------------
' REVISIONS
'   NCJ 28 Nov 02 - Added file header; removed AREZZO Settings form
'   ASH 2/12/2002 - Added correct enabling and disabling of mnuFUnregisterdatabase
'                   also correctly refreshed treeview and databaseinfo form after
'                   unregistering of database
'   TA 20/12/2002 - Reinstated old status bar as used in other modules
'   NCJ 23 Apr 03 - Use generic MACRO Help for 3.0
'   REM 12/03/04 - Added condition compilation arguments in SetMenuItems and CheckUserRights routines so
'                  certain menu items are not displayed in the Desktop edition of MACRO
'   REM 15/07/04 - Added check for 'Maintain User Role' permission on the 'New Role..' menu item
'   MLM 21/06/05: bug 2528: Diasable menus until application is initialised.
'   ic 07/12/2005 added active directory servers menu enablement
'   NCJ 28 Feb 08 - Bug 3003 - Enable/disable Security DB menu on "Register/Create DB" permission
'------------------------------------------------------------------------------

Option Explicit

Private WithEvents mofrmTreeView As frmSysAdminTreeView
Attribute mofrmTreeView.VB_VarHelpID = -1
Private WithEvents mofrmMain As Form
Attribute mofrmMain.VB_VarHelpID = -1
'Private WithEvents mofrmMain As frmMenu

'store popup item selected
Private mlPopUpItem As Long

Private msDatabaseCode As String
Private msStudyName As String
Private msSiteCode As String
Private mlStudyId As Long
Private msUserName As String
Private msRoleCode As String
Private menNodeTag As eSMNodeTag


'---------------------------------------------------------------------
Private Sub MDIForm_Load()
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    FormCentre Me
    
    'mnuVUnusableRoleAssignments.Checked = mnuVUnusableRoleAssignments
    
End Sub

'---------------------------------------------------------------------
Private Sub MDIForm_Resize()
'---------------------------------------------------------------------
Dim lMinWidth As Long
Dim lMinHeight As Long
Dim lProposedHeight As Long
Dim lProposedWidth As Long

    'detect a minimize call and exit the rest of Resize
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    
    'used to avoid flickering by avoiding re-calculation of height and width
    lProposedWidth = Me.Width
    lProposedHeight = Me.Height
    
    If Not mofrmTreeView Is Nothing Then
        mofrmTreeView.Height = Me.ScaleHeight
        'do not resize below the width of treeview
        If Me.ScaleWidth < mofrmTreeView.Width Then
            lProposedWidth = Me.Width + mofrmTreeView.Width - Me.ScaleWidth
        End If
        
        If Not mofrmMain Is Nothing Then
        'special case for rolemanagemnet form
            If mofrmMain.Name = "frmRoleManagement" Then
                'minimum width of the role form
                lMinWidth = mofrmMain.txtRoleDescription.Width + 2000
                '1800 = heights of 4 command buttons plus arbitrary figure
                'for spaces between the buttons
                lMinHeight = 1800 + mofrmMain.cmdRemoveOne.Top
                If Me.ScaleWidth < mofrmTreeView.Width + lMinWidth Then
                    lProposedWidth = Me.Width + mofrmTreeView.Width - Me.ScaleWidth + lMinWidth
                End If
                If Me.ScaleHeight < lMinHeight Then
                  lProposedHeight = Me.Height + lMinHeight - Me.ScaleHeight
                End If
            End If
            'reset height and width of form main
            mofrmMain.Height = lProposedHeight - Me.Height + Me.ScaleHeight
            mofrmMain.Width = lProposedWidth - Me.Width + Me.ScaleWidth - mofrmTreeView.Width
        End If
        mofrmTreeView.Height = lProposedHeight - Me.Height + Me.ScaleHeight
    End If
    
    If Me.WindowState = vbNormal Then
        Me.Width = lProposedWidth
        Me.Height = lProposedHeight
    End If
End Sub

'---------------------------------------------------------------------
Public Sub CreateDatabaseForm()
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------

    frmNewDatabase.Show vbModal

End Sub

#If rochepatch <> 1 Then
'---------------------------------------------------------------------
Private Sub mnuActiveDirectoryServers_Click()
'---------------------------------------------------------------------
' ic 05/12/2005
' active directory servers menu option
'---------------------------------------------------------------------
    Call ActiveDirectoryServers
End Sub

'---------------------------------------------------------------------
Public Sub ActiveDirectoryServers()
'---------------------------------------------------------------------
' ic 05/12/2005
' display active directory servers config window
'---------------------------------------------------------------------
    Call frmActiveDirectoryServers.Show(vbModal)
End Sub
#End If
'---------------------------------------------------------------------
Private Sub mnuDataComm_Click()
'---------------------------------------------------------------------

    Call DataCommunication

End Sub

'---------------------------------------------------------------------
Public Sub DataCommunication()
'---------------------------------------------------------------------

    Call frmDataCommunication.Show(vbModal)

End Sub

'---------------------------------------------------------------------
Private Sub mnuDCreateDatabase_Click()
'---------------------------------------------------------------------

    Call CreateDatabaseForm

End Sub

'---------------------------------------------------------------------
Public Sub LockAdministrationForm()
'---------------------------------------------------------------------
' MLM 11/06/03: Launch Database Lock Admin form for the selected db, rather than the
'   logged into db.
'---------------------------------------------------------------------

    Call frmLocksAdmin.Display(goUser, msDatabaseCode)
    'Call frmLocksAdmin.Display(goUser, goUser.DatabaseCode)

End Sub

'--------------------------------------------------------------------
Private Sub mnuDistribute_Click()
'--------------------------------------------------------------------

      Call frmExportLab.Display(True, goUser.Database.DatabaseCode)

End Sub

'---------------------------------------------------------------------
Private Sub mnuDLockAdministration_Click()
'---------------------------------------------------------------------

     Call LockAdministrationForm

End Sub

'---------------------------------------------------------------------
Public Sub RegisterDatabaseForm()
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------

    frmRegisterDatabase.Show vbModal
    Me.RefereshTreeView

End Sub

'---------------------------------------------------------------------
Private Sub mnuDRegisterDatabase_Click()
'---------------------------------------------------------------------

    Call RegisterDatabaseForm
    
End Sub

'---------------------------------------------------------------------
Public Sub RestoreSiteDatabaseForm()
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------

    frmRestoreSiteDatabase.Show vbModal

End Sub

'---------------------------------------------------------------------
Private Sub mnuDRestoreSiteDatabase_Click()
'---------------------------------------------------------------------

    Call RestoreSiteDatabaseForm

End Sub

'---------------------------------------------------------------------
Public Sub SetDatabasePasswordForm(sDatabaseCode As String)
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------

    Call frmChangeDatabasePassword.Display(sDatabaseCode)
    RefereshTreeView

End Sub

'---------------------------------------------------------------------
Private Sub mnuDSecurityDatabase_Click()
'---------------------------------------------------------------------

    Call SecurityDBForm

End Sub

'---------------------------------------------------------------------
Public Sub SecurityDBForm()
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------

    Call frmSecurityDB.Display

End Sub

'---------------------------------------------------------------------
Private Sub mnuDSetDatabasePassword_Click()
'---------------------------------------------------------------------

    Call SetDatabasePasswordForm(msDatabaseCode)


End Sub

'--------------------------------------------------------------------
Public Sub mnuDTimezone_Click()
'--------------------------------------------------------------------
' MLM 09/06/03: Added. Show a new form to display/set the db's timezone.
'--------------------------------------------------------------------

    frmDatabaseTimezone.Display msDatabaseCode

End Sub

'--------------------------------------------------------------------
Private Sub mnuDunRegisterDatabase_Click()
'--------------------------------------------------------------------
'
'--------------------------------------------------------------------

     Call UnRegisterDatabase(msDatabaseCode)

End Sub



'------------------------------------------------------------------
Private Sub mnuExport_Click()
'------------------------------------------------------------------

    Call frmExportLab.Display(False, goUser.Database.DatabaseCode)

End Sub

'-------------------------------------------------------------------
Private Sub mnuExportStudyDefinition_Click()
'-------------------------------------------------------------------

    frmExportStudyDefinition.Display

End Sub

'---------------------------------------------------------------------
Private Sub mnuGenerateHTML_Click()
'---------------------------------------------------------------------

    frmGenerateHTML.Display

End Sub

'---------------------------------------------------------------------
Private Sub mnuHAbout_Click()
'---------------------------------------------------------------------

    frmAbout.Display
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuHTMLFolder_Click()
'---------------------------------------------------------------------

    Call ChangeHTMLFolder

End Sub

'---------------------------------------------------------------------
Private Sub mnuHUserGuide_Click()
'---------------------------------------------------------------------

    Call MACROHelp(Me.hWnd, App.Title)

End Sub

'------------------------------------------------------------------
Private Sub mnuImport_Click()
'------------------------------------------------------------------

    Call frmImportPatientData.Display(goUser.Database.DatabaseCode, True)

End Sub

'-------------------------------------------------------------------
Private Sub mnuImportStudyDefinition_Click()
'-------------------------------------------------------------------
    
    frmImportStudyDefinition.Display
    Me.RefereshTreeView

End Sub

'--------------------------------------------------------------------
Private Sub mnuLaboratorySiteAdministration_Click()
'--------------------------------------------------------------------

    Call frmTrialSiteAdmin.Display(goUser.Database.DatabaseCode, DisplaySitesByLab)

End Sub

'---------------------------------------------------------------------
Private Sub mnuLoginLog_Click()
'---------------------------------------------------------------------
'Display the Login Log details
'---------------------------------------------------------------------

    Call LogDetailsForm(False)
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuLogOff_Click()
'---------------------------------------------------------------------

    Call UserLogOff(False)

End Sub

'---------------------------------------------------------------------
Public Function DatabaseTagInfo() As Collection
'---------------------------------------------------------------------

    Set DatabaseTagInfo = mofrmTreeView.DatabaseTags

End Function

'---------------------------------------------------------------------
Private Sub ChangeHTMLFolder()
'---------------------------------------------------------------------
Dim sHTMLLocation As String
Dim sSecHTMLLocation As String
Dim sDatabaseCode As String
Dim sSQL As String
Dim bUpdateHTMLLoaction As Boolean

    sHTMLLocation = goUser.Database.HTMLLocation
    sSecHTMLLocation = goUser.Database.SecureHTMLLocation
    sDatabaseCode = goUser.Database.DatabaseCode

    Call frmSetHTMLFolder.HTMLPath(sHTMLLocation, sSecHTMLLocation)
    
    bUpdateHTMLLoaction = goUser.Database.UpdateHTMLLoaction(SecurityADODBConnection, sDatabaseCode, sHTMLLocation, sSecHTMLLocation)
    
    RefereshDatabaseInfoForm
    
End Sub

'---------------------------------------------------------------------
Public Sub RoleManagementForm(Optional sRoleCode As String)
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------

    'display role editor
    If mofrmMain.Name <> "frmRoleManagement" Then
        Unload mofrmMain
        Set mofrmMain = New frmRoleManagement
    End If
    mofrmMain.Display sRoleCode
    
    sRoleCode = ""
    msUserName = ""
    msRoleCode = sRoleCode
    Call CheckUserRights
    
    'in case a new child form was shown, resize it to fit the MDI form
    mofrmMain.Top = 0
    mofrmMain.Left = mofrmTreeView.Width
    Call MDIForm_Resize

End Sub

'--------------------------------------------------------------------
Private Sub mnuRGoToRole_Click()
'--------------------------------------------------------------------

    Call RoleManagementForm(msRoleCode)

End Sub

'---------------------------------------------------------------------
Private Sub mnuRNewRole_Click()
'---------------------------------------------------------------------

    Call RoleManagementForm
 
End Sub

'--------------------------------------------------------------------
Private Sub mnuSiteAdministration_Click()
'--------------------------------------------------------------------
        
    Call FrmSiteAdmin.Display(goUser.Database.DatabaseCode, msStudyName, msSiteCode)
    RefereshTreeView

End Sub

'--------------------------------------------------------------------
Private Sub mnuStudySiteAdministration_Click()
'--------------------------------------------------------------------

    Call frmTrialSiteAdminVersioning.Display(goUser.Database.DatabaseCode, DisplaySitesByTrial, "")
    RefereshTreeView
    
End Sub

'--------------------------------------------------------------------
Private Sub mnuStudyStatus_Click()
'--------------------------------------------------------------------

    frmTrialStatus.Display (goUser.Database.DatabaseCode)

End Sub

'--------------------------------------------------------------------
Private Sub mnuSuExport_Click()
'--------------------------------------------------------------------

    Call frmExportPatientData.Display(goUser.Database.DatabaseCode)

End Sub

'--------------------------------------------------------------------
Private Sub mnuSuImport_Click()
'--------------------------------------------------------------------

    frmImportPatientData.Display (goUser.Database.DatabaseCode)

End Sub

'---------------------------------------------------------------------
Private Sub mnuUDeleteUserRole_Click()
'---------------------------------------------------------------------
Dim sMSG As String
Dim sRoleToDelete As String
    
    sRoleToDelete = msStudyName & "|" & msSiteCode & "|" & msRoleCode & "|" & msUserName
    sMSG = "Are you sure you want to delete the selected user role?"
    If DialogQuestion(sMSG) = vbYes Then
        Call frmNewUserRole.DeleteRoles(sRoleToDelete, msDatabaseCode, True)
        Me.RefereshUserRoleInfoForm
    End If

End Sub

'---------------------------------------------------------------------
Private Sub mnuUExit_Click()
'---------------------------------------------------------------------

    Call UserLogOff(True)

End Sub

'---------------------------------------------------------------------
Public Sub UserRoleForm(Optional bNew As Boolean = True, Optional sUsername As String = "", _
                        Optional sDatabaseCode As String = "", Optional sRoleCode As String = "", _
                        Optional sStudyCode As String = "", Optional sSiteCode As String)
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------

    Call frmNewUserRole.Display(bNew, sUsername, sDatabaseCode, sRoleCode, sStudyCode, sSiteCode)

End Sub

'---------------------------------------------------------------------
Public Sub NewUserForm(Optional bNewUser As Boolean = False, _
                        Optional sUser As String = "")
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------

    'display user properties
    If mofrmMain.Name <> "frmUserDetails" Then
        Unload mofrmMain
        Set mofrmMain = New frmUserDetails
    End If
    If msUserName = "" Then
        msUserName = sUser
    End If
    
    mofrmMain.Display bNewUser, msUserName
    msUserName = ""
    msRoleCode = ""
    Call CheckUserRights
    'in case a new child form was shown, resize it to fit the MDI form
    mofrmMain.Top = 0
    mofrmMain.Left = mofrmTreeView.Width
    Call MDIForm_Resize


End Sub

'---------------------------------------------------------------------
Private Sub mnuUGoToUser_Click()
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------

    Call NewUserForm(False)

End Sub

'---------------------------------------------------------------------
Private Sub mnuUNewUser_Click()
'---------------------------------------------------------------------

    Call NewUserForm(True)

End Sub

'---------------------------------------------------------------------
Private Sub mnuUNewUserRole_Click()
'---------------------------------------------------------------------
    
    Call UserRoleForm

End Sub

'---------------------------------------------------------------------
Private Sub mnuUPasswordPolicy_Click()
'---------------------------------------------------------------------

    Call frmPasswordPolicy.Display

End Sub

'---------------------------------------------------------------------
Private Sub mnuUpgradeDatabase_Click()
'---------------------------------------------------------------------
'REM 11/06/03 - upgrade a MACRO database to latest 3.0 version
'---------------------------------------------------------------------
    
    Call UpgradeDatabase

End Sub

'---------------------------------------------------------------------
Public Sub UpgradeDatabase()
'---------------------------------------------------------------------
'Upgrade a selected database
'---------------------------------------------------------------------
Dim sMessage As String
Dim sDBVersion As String

    sDBVersion = mofrmTreeView.DatabaseVersion

    Call UpgradeToLatestMACRODatabase(SecurityADODBConnection.ConnectionString, msDatabaseCode, sDBVersion, sMessage)
    
    Call RefereshTreeView
    
    Call RefereshDatabaseInfoForm

End Sub

'---------------------------------------------------------------------
Private Sub mnuUResetPassword_Click()
'---------------------------------------------------------------------
' Allow the logged in user to change their own password.
'---------------------------------------------------------------------

    'TA 7/6/2000 SR359:  updated so display function is called)
    'REM 10/10/02 - Changed to new form
    If frmPasswordChange.Display(goUser, gsSecCon) Then
        Call goUser.gLog(goUser.UserName, gsCHANGE_PSWD, "System administrator changed password")
    End If

End Sub

'---------------------------------------------------------------------
Public Sub RefereshTreeView()
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------

    Call mofrmTreeView.RefreshTreeView
    
    Call RestartSystemIdleTimer

End Sub

'---------------------------------------------------------------------
Public Sub RefereshDatabaseInfoForm()
'---------------------------------------------------------------------

    'display user properties
    If mofrmMain.Name = "frmDatabaseInfo" Then
        mofrmMain.RefreshDatabases
    End If
    
End Sub

'---------------------------------------------------------------------
Public Sub RefereshUserRoleInfoForm()
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------

    If mofrmMain.Name <> "frmUserRolesInfo" Then
        Unload mofrmMain
        Set mofrmMain = New frmUserRolesInfo
    End If
    mofrmMain.Display msDatabaseCode, msStudyName, msSiteCode, menNodeTag

End Sub


'---------------------------------------------------------------------
Private Sub mnuUSystemPolicy_Click()
'---------------------------------------------------------------------

    frmSystemProperties.Show vbModal

End Sub

'---------------------------------------------------------------------
Private Sub mnuVCommunicationLog_Click()
'---------------------------------------------------------------------

    Call frmCommunicationsLog.Display

End Sub

'---------------------------------------------------------------------
Private Sub mnuVRefresh_Click()
'---------------------------------------------------------------------
'TA 20/01/2006: force expand when refreshing from menu
'---------------------------------------------------------------------

    Call mofrmTreeView.RefreshTreeView(True)
    
    Call RestartSystemIdleTimer

End Sub

'---------------------------------------------------------------------
Private Sub mnuVSystemLog_Click()
'---------------------------------------------------------------------
'Display the system log form
'---------------------------------------------------------------------

   Call LogDetailsForm(True)

End Sub

'---------------------------------------------------------------------
Public Sub LogDetailsForm(bLogDetails As Boolean, Optional sDatabaseCode As String = "")
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------

    Call frmLogDetails.Display(bLogDetails, sDatabaseCode)

End Sub

'---------------------------------------------------------------------------------------
Public Sub SelectedItemParameters(enNodeTag As eSMNodeTag, sDatabaseCode As String, _
                                    lStudyId As Long, sStudyName As String, _
                                    sSiteCode As String, sUsername As String, _
                                    sRoleCode As String)
'----------------------------------------------------------------------------------------
'assigns value to module level variables when an item is clicked in the right hand pane
'----------------------------------------------------------------------------------------

    msDatabaseCode = sDatabaseCode
    msStudyName = sStudyName
    msSiteCode = sSiteCode
    mlStudyId = lStudyId
    msUserName = sUsername
    msRoleCode = sRoleCode

    Call CheckUserRights
    
    Call SetMenuItems(enNodeTag)
    
End Sub

'---------------------------------------------------------------------------------------
Private Sub SetMenuItems(enNodeTag As eSMNodeTag)
'---------------------------------------------------------------------------------------
'REM 12/03/04 - added conditional comp arguments so certain menu items are disabled in MACRO Desktop Ed
'---------------------------------------------------------------------------------------

    Select Case enNodeTag
    Case eSMNodeTag.DatabaseTag
        mnuDLockAdministration.Enabled = True
        mnuDSetDatabasePassword.Enabled = True
        mnuLoginLog.Enabled = True
        mnuDTimezone.Enabled = True
        mnuDunRegisterDatabase.Enabled = (goUser.DatabaseCode <> msDatabaseCode)

    Case eSMNodeTag.DisconnectedDB
        mnuDLockAdministration.Enabled = False
        mnuDSetDatabasePassword.Enabled = True
        mnuLoginLog.Enabled = False
        mnuDTimezone.Enabled = False
        mnuUpgradeDatabase.Enabled = False
    Case eSMNodeTag.Upgrade
        mnuDLockAdministration.Enabled = False
        mnuDSetDatabasePassword.Enabled = False
        mnuLoginLog.Enabled = False
        mnuDTimezone.Enabled = False
        mnuUpgradeDatabase.Enabled = True
    End Select

    'If using MACRO Desktop then always disable the following menu items regardless of user permissions
    #If DESKTOP = 1 Then
        mnuDCreateDatabase.Enabled = False
        mnuDRegisterDatabase.Enabled = False
        mnuDunRegisterDatabase.Enabled = False
        mnuDSetDatabasePassword.Enabled = False
        mnuHTMLFolder.Enabled = False
        mnuDSecurityDatabase = False
        mnuDRestoreSiteDatabase.Enabled = False
    #End If


End Sub

'---------------------------------------------------------------------------------------
Private Sub mofrmMain_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------------------------

    Call RestartSystemIdleTimer

End Sub

'---------------------------------------------------------------------
Private Sub mofrmMain_Load()
'---------------------------------------------------------------------
'MDIForm_Resize
'---------------------------------------------------------------------
    
    mofrmMain.Visible = False
    
    Call RestartSystemIdleTimer
    
End Sub

'---------------------------------------------------------------------
Public Sub MDIResize()
'---------------------------------------------------------------------
'ASH 3/12/2002 Commented out as per request from Matthew M.
'---------------------------------------------------------------------
    
    'Call MDIForm_Resize

End Sub

'---------------------------------------------------------------------
Private Sub mofrmTreeView_Resize(ByRef sglWidth As Single)
'---------------------------------------------------------------------
' Called when the divider between the L and R sides is moved
'---------------------------------------------------------------------
    
    If sglWidth > Me.ScaleWidth Then
        sglWidth = Me.ScaleWidth
    End If
    'Debug.Print sglWidth
    mofrmMain.Left = sglWidth
    mofrmMain.Width = Me.ScaleWidth - sglWidth
    
End Sub

'---------------------------------------------------------------------
Private Sub MDIForm_Unload(Cancel As Integer)
'---------------------------------------------------------------------

'---------------------------------------------------------------------
    Call UserLogOff(True)
    Call ExitMACRO

End Sub

'---------------------------------------------------------------------
Public Property Get Mode() As String
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    Mode = gsSYSTEM_MANAGER_MODE

End Property

'---------------------------------------------------------------------
 Public Sub CheckUserRights()
'---------------------------------------------------------------------
'ASH 2/12/2002
'REM 12/03/04 - added conditional comp arguments so certain menu items are disabled in MACRO Desktop Ed
'REM 15/07/04 - Added check for 'Maintain User Role' permission on the 'New Role..' menu item
'ic 07/12/2005 added active directory servers menu enablement
'---------------------------------------------------------------------
#If rochepatch <> 1 Then
    'active directory servers
    If (gbActiveDirectory) Then
        If goUser.CheckPermission(gsFnActiveDirectoryServers) Then
            mnuActiveDirectoryServers.Enabled = True
        Else
             mnuActiveDirectoryServers.Enabled = False
        End If
        If (gbIsActiveDirectoryLogin) Then
            'cant log out and in as another user with active directory
            mnuLogOff.Enabled = False
        End If
    Else
        mnuActiveDirectoryServers.Enabled = False
        mnuLogOff.Enabled = True
    End If
#Else
    mnuActiveDirectoryServers.Visible = False
#End If

    'password policy and system properties
    If goUser.CheckPermission(gsFnChangeSystemProperties) Then
        mnuUPasswordPolicy.Enabled = True
        mnuUSystemPolicy.Enabled = True
    Else
         mnuUPasswordPolicy.Enabled = False
         mnuUSystemPolicy.Enabled = False
    End If
    'create database
    If goUser.CheckPermission(gsFnCreateDB) Then
        mnuDCreateDatabase.Enabled = True
    Else
         mnuDCreateDatabase.Enabled = False
    End If
    
    ' NCJ 28 Feb 08 - Bug 3003 - Also enable/disable Security DB menu item
    mnuDSecurityDatabase.Enabled = _
        (goUser.CheckPermission(gsFnRegisterDB) Or goUser.CheckPermission(gsFnCreateDB))
        
    'register database
    If goUser.CheckPermission(gsFnRegisterDB) Then
        mnuDRegisterDatabase.Enabled = True
        'MLM 09/06/03:
        mnuDTimezone.Enabled = (msDatabaseCode <> "")
    Else
        mnuDRegisterDatabase.Enabled = False
    End If
    'unregister database
    If goUser.CheckPermission(gsFnUnRegisterDatabase) And (msDatabaseCode <> "") Then
        mnuDunRegisterDatabase.Enabled = True
    Else
         mnuDunRegisterDatabase.Enabled = False
    End If
    'restore site database
    If goUser.CheckPermission(gsFnRestoreDatabase) Then
        mnuDRestoreSiteDatabase.Enabled = True
    Else
         mnuDRestoreSiteDatabase.Enabled = False
    End If
    'set database password
    'MLM 11/06/03: You can only change a database password when a database is selected..
    If goUser.CheckPermission(gsFnChangePassword) And (msDatabaseCode <> "") Then
        mnuDSetDatabasePassword.Enabled = True
    Else
         mnuDSetDatabasePassword.Enabled = False
    End If
    'lock administration
    If (goUser.CheckPermission(gsFnRemoveOwnLocks) Or goUser.CheckPermission(gsFnRemoveAllLocks)) _
        And (msDatabaseCode <> "") Then
        mnuDLockAdministration.Enabled = True
    Else
         mnuDLockAdministration.Enabled = False
    End If
    'new user
    If goUser.CheckPermission(gsFnCreateNewUser) Then
        mnuUNewUser.Enabled = True
    Else
         mnuUNewUser.Enabled = False
    End If
    'go to user
    If goUser.CheckPermission(gsFnCreateNewUser) And (msUserName <> "") Then
        mnuUGoToUser.Enabled = True
    Else
         mnuUGoToUser.Enabled = False
    End If
    'New user role
    If goUser.CheckPermission(gsFnMaintRole) Then
        mnuUNewUserRole.Enabled = True
    Else
         mnuUNewUserRole.Enabled = False
    End If
    'REM 15/07/04 - Added check for create new role
    'New role
    If goUser.CheckPermission(gsFnMaintRole) Then
        mnuRNewRole.Enabled = True
    Else
        mnuRNewRole.Enabled = False
    End If
    'go to role
    If goUser.CheckPermission(gsFnMaintRole) And (msRoleCode <> "") Then
        mnuRGoToRole.Enabled = True
    Else
         mnuRGoToRole.Enabled = False
    End If
    'delete user role
    If goUser.CheckPermission(gsFnMaintRole) And (msUserName <> "") Then
        mnuUDeleteUserRole.Enabled = True
    Else
         mnuUDeleteUserRole.Enabled = False
    End If
    'reset my password
    If goUser.CheckPermission(gsFnResetPassword) Then
        mnuUResetPassword.Enabled = True
    Else
         mnuUResetPassword.Enabled = False
    End If
    'communications log
    If goUser.CheckPermission(gsFnViewSiteServerCommunication) Then
        mnuVCommunicationLog.Enabled = True
    Else
         mnuVCommunicationLog.Enabled = False
    End If
    'system log
    If goUser.CheckPermission(gsFnViewSystemLog) Then
        mnuVSystemLog.Enabled = True
    Else
        mnuVSystemLog.Enabled = False
    End If
    'user log
    If goUser.CheckPermission(gsFnViewSystemLog) Then
        mnuLoginLog.Enabled = True
    Else
         mnuLoginLog.Enabled = False
    End If
    
    'site/study/laboratory administration
    If goUser.CheckPermission(gsFnCreateSite) Then
        mnuSiteAdministration.Enabled = True
        mnuStudySiteAdministration.Enabled = True
        mnuLaboratorySiteAdministration.Enabled = True
    Else
        mnuSiteAdministration.Enabled = False
        mnuStudySiteAdministration.Enabled = False
        mnuLaboratorySiteAdministration.Enabled = False
    End If
        
    'export study
    If goUser.CheckPermission(gsFnImportStudyDef) Then
        mnuExportStudyDefinition.Enabled = True
    Else
         mnuExportStudyDefinition.Enabled = False
    End If
    'import study
    If goUser.CheckPermission(gsFnImportStudyDef) Then
        mnuImportStudyDefinition.Enabled = True
    Else
         mnuImportStudyDefinition.Enabled = False
    End If
    'study status
    If goUser.CheckPermission(gsFnChangeTrialStatus) Then
        mnuStudyStatus.Enabled = True
    Else
         mnuStudyStatus.Enabled = False
    End If
    'lab export,import,distribute
    If goUser.CheckPermission(gsFnMaintainLaboratories) Then
        mnuExport.Enabled = True
        mnuImport.Enabled = True
        mnuDistribute.Enabled = True
    Else
        mnuExport.Enabled = False
        mnuImport.Enabled = False
        mnuDistribute.Enabled = False
    End If
    'export subject
    If goUser.CheckPermission(gsFnExportPatData) Then
        mnuSuExport.Enabled = True
    Else
         mnuSuExport.Enabled = False
    End If
    'import subject
    If goUser.CheckPermission(gsFnImportPatData) Then
        mnuSuImport.Enabled = True
    Else
         mnuSuImport.Enabled = False
    End If
    
    'If using MACRO Desktop then always disable the following menu items regardless of user permissions
    #If DESKTOP = 1 Then
        mnuDCreateDatabase.Enabled = False
        mnuDRegisterDatabase.Enabled = False
        mnuDunRegisterDatabase.Enabled = False
        mnuDSetDatabasePassword.Enabled = False
        mnuHTMLFolder.Enabled = False
        mnuDSecurityDatabase = False
        mnuDRestoreSiteDatabase.Enabled = False
    #End If
    
End Sub

'---------------------------------------------------------------------
Public Sub InitialiseMe()
'---------------------------------------------------------------------
' Perform initialisations specific to System Management
' Called from Main in MainMacroModule
' NCJ 16 Sept 1999
'---------------------------------------------------------------------


    'TA 20/12/2002: show menu for current db
    'mnuCurrentDB.Caption = goUser.DatabaseCode & "-&Tasks"
    'mnuCurrentDB.Visible = True
    
    ' NCJ 13/1/00 - Set System Management user guide
    ' NCJ 23 Apr 03 - Use generic MACRO Help for 3.0
'    gsMACROUserGuidePath = gsMACROUserGuidePath & "SM\Contents.htm"
   
    'REM 02/04/03 - Disable this menu item to start with, as it was allowing users to unregister the database they were logged into.
                    'It will become enabled when a user clicks on a database they are allowed to unregister
    mnuDunRegisterDatabase.Enabled = False
    
    mnuUpgradeDatabase.Enabled = False
        
    Set mofrmTreeView = New frmSysAdminTreeView
    mofrmTreeView.Show
'    frmMain.Show
    mofrmTreeView.Top = 0
    mofrmTreeView.Left = 0
'    frmMain.Top = 0
    
    'MLM 17/10/02: Default to showing the databases list in the right-hand pane.
    Set mofrmMain = New frmDatabaseInfo
    mofrmMain.Display
    mofrmMain.Top = 0
    mofrmMain.Left = mofrmTreeView.Width
    
    Call MDIForm_Resize

    'MLM 20/06/05: bug 2500: enable all the top-level menus only after successful login..
    EnableMenus
    'ash 4/12/2002
    Call CheckUserRights
    
End Sub

'---------------------------------------------------------------------
Private Sub EnableMenus()
'---------------------------------------------------------------------
' MLM 20/06/05: bug 2500: Created
'---------------------------------------------------------------------
    mnuFile.Enabled = True
    mnuDatabase.Enabled = True
    mnuUser.Enabled = True
    mnuRole.Enabled = True
    mnuView.Enabled = True
    mnuCurrentDB.Enabled = True
    mnuHelp.Enabled = True
End Sub

'---------------------------------------------------------------------
Private Sub mofrmTreeView_SelectedNode(enNodeTag As eSMNodeTag, sDatabaseCode As String, lStudyId As Long, sStudyName As String, sSiteCode As String, sUsername As String, sRoleCode As String)
'---------------------------------------------------------------------
' MLM 17/10/02: Created. Remember what is selected in the treeview, for use in menu items
'   Display the appropriate form in the right hand side.
'---------------------------------------------------------------------

    'store selected item
    msDatabaseCode = sDatabaseCode
    msStudyName = sStudyName
    mlStudyId = lStudyId
    msSiteCode = sSiteCode
    msUserName = ""
    msRoleCode = ""
    menNodeTag = enNodeTag
    
    'load main form
    Select Case enNodeTag
    Case eSMNodeTag.DatabasesTag
        'display databases list
        If mofrmMain.Name <> "frmDatabaseInfo" Then
            Unload mofrmMain
            Set mofrmMain = New frmDatabaseInfo
            mofrmMain.Display
        End If
        
    Case eSMNodeTag.DatabaseTag, eSMNodeTag.SitesTag, eSMNodeTag.SiteTag, eSMNodeTag.StudiesTag, eSMNodeTag.StudyTag, eSMNodeTag.DisconnectedDB
        'display a list of relevant user roles
        If mofrmMain.Name <> "frmUserRolesInfo" Then
            Unload mofrmMain
            Set mofrmMain = New frmUserRolesInfo
        End If
        mofrmMain.Display sDatabaseCode, sStudyName, sSiteCode, enNodeTag
        
    Case eSMNodeTag.RoleTag
        'display role editor
        If mofrmMain.Name <> "frmRoleManagement" Then
            Unload mofrmMain
            Set mofrmMain = New frmRoleManagement
        End If
        mofrmMain.Display sRoleCode
        
    Case eSMNodeTag.UserTag
        'display user properties
        If mofrmMain.Name <> "frmUserDetails" Then
            Unload mofrmMain
            Set mofrmMain = New frmUserDetails
        End If
        mofrmMain.Display False, sUsername
        
    Case Else
        'clicking on other nodes has no effect on the right hand side
    End Select
    
    Call CheckUserRights
    
    'in case a new child form was shown, resize it to fit the MDI form
    mofrmMain.Top = 0
    mofrmMain.Left = mofrmTreeView.Width
    Call MDIForm_Resize

End Sub

'---------------------------------------------------------------------
Private Sub mnuVUnusableRoleAssignments_Click()
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------

    If mofrmTreeView.ViewUnusableRoles = True Then
        mofrmTreeView.ViewUnusableRoles = False
        mnuVUnusableRoleAssignments.Checked = False
        mofrmTreeView.RefreshTreeView
    Else
        mofrmTreeView.ViewUnusableRoles = True
        mnuVUnusableRoleAssignments.Checked = True
        mofrmTreeView.RefreshTreeView
    End If

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
        If frmTimeOutSplash.Display(False) Then
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
Private Sub mnuPopUpSubItem_Click(Index As Integer)
'---------------------------------------------------------------------
'store item clicked on user-defined menu
'---------------------------------------------------------------------
     mlPopUpItem = Index + 1
End Sub

'---------------------------------------------------------------------
Public Function ShowPopUpMenu(oMenuItems As clsMenuItems) As String
'---------------------------------------------------------------------
' Show a user-defined popup menu
' Input: collection of clsMenuItem objects
'
' Output:
'       function - key item selected ("" if nothing selected)
'---------------------------------------------------------------------
Dim i As Long

    If oMenuItems.Count = 0 Then Exit Function
    
    For i = 0 To oMenuItems.Count - 1
        If i <> 0 Then
            Load Me.mnuPopUpSubItem(i)
        End If
        With mnuPopUpSubItem(i)
            .Enabled = oMenuItems.Item(i).Enabled
            .Checked = oMenuItems.Item(i).Checked
            .Caption = oMenuItems.Item(i).Caption
        End With
    Next
    
    'set default choice to unspecified
    mlPopUpItem = -1
    
    'show menu
    If oMenuItems.DefaultItemIndex = -1 Then
        'show with no default
        PopupMenu mnuPopUp
    Else
        PopupMenu mnuPopUp, , , , mnuPopUpSubItem(oMenuItems.DefaultItemIndex)
    End If
    
    'unload controls created at run time (except 0 element)
    For i = 1 To mnuPopUpSubItem.Count - 1
        Unload mnuPopUpSubItem(i)
    Next
    
    'return user's choice
    If mlPopUpItem = -1 Then
        ShowPopUpMenu = ""
    Else
        ShowPopUpMenu = oMenuItems.Item(mlPopUpItem - 1).Key
    End If
    
End Function


'------------------------------------------------------------------------------------
Public Sub UnRegisterDatabase(ByVal sDatabaseCode As String)
'------------------------------------------------------------------------------------
'ash 18/11/2002
'Deregisters/Unregisters a selected database
'------------------------------------------------------------------------------------
Dim sSQL As String
Dim rsDatabases As ADODB.Recordset
Dim sMSG As String
Dim i As Integer

    On Error GoTo ErrHandler
    
    If sDatabaseCode = "" Then Exit Sub
    
    'if its the last database then do not delete
    sSQL = "SELECT * FROM Databases"
    Set rsDatabases = New ADODB.Recordset
    rsDatabases.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
    If rsDatabases.RecordCount = 1 Then
        sMSG = "The last database cannot be unregistered"
        Call DialogInformation(sMSG)
        Exit Sub
    End If

    sMSG = "Unregistering " & sDatabaseCode & " will make it unusable until re-registered." & vbCrLf & "Are you sure you want to unregister it?"
    If DialogQuestion(sMSG, "Unregister Database") = vbYes Then
        sSQL = "Delete FROM UserDatabase WHERE DatabaseCode = '" & sDatabaseCode & "'"
        SecurityADODBConnection.Execute sSQL
        sSQL = "Delete FROM Databases WHERE DatabaseCode = '" & sDatabaseCode & "'"
        SecurityADODBConnection.Execute sSQL
            
        Me.RefereshDatabaseInfoForm
        Me.RefereshTreeView
        'select the first database node
        
    
    End If

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.UnRegisterDatabase"
End Sub

'---------------------------------------------------------------------
Public Function ForgottenPassword(sSecurity As String, sUsername As String, sPassword As String, sErrMsg As String) As eDTForgottenPassword
'---------------------------------------------------------------------
'REM 06/12/02
'---------------------------------------------------------------------

    'Dummy routine

End Function

