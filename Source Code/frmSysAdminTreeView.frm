VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSysAdminTreeView 
   BorderStyle     =   0  'None
   Caption         =   "User Management"
   ClientHeight    =   8220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4320
   Icon            =   "frmsysAdminTreeView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8220
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSep 
      Height          =   7815
      Left            =   3960
      MousePointer    =   9  'Size W E
      TabIndex        =   1
      Top             =   120
      Width           =   135
   End
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   2760
      Top             =   7320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsysAdminTreeView.frx":0442
            Key             =   "users"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsysAdminTreeView.frx":059C
            Key             =   "databaseloggedin"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsysAdminTreeView.frx":0B36
            Key             =   "usersys"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsysAdminTreeView.frx":0F88
            Key             =   "databasedisconnected"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsysAdminTreeView.frx":1522
            Key             =   "macro2"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsysAdminTreeView.frx":1974
            Key             =   "databases"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsysAdminTreeView.frx":1F0E
            Key             =   "role"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsysAdminTreeView.frx":2068
            Key             =   "roles"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsysAdminTreeView.frx":21C2
            Key             =   "inactivesite"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsysAdminTreeView.frx":231C
            Key             =   "activesitewarning"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsysAdminTreeView.frx":2476
            Key             =   "activesite"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsysAdminTreeView.frx":25D0
            Key             =   "sites"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsysAdminTreeView.frx":272A
            Key             =   "studies"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsysAdminTreeView.frx":2884
            Key             =   "studyclosedwarning"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsysAdminTreeView.frx":29DE
            Key             =   "studyclosed"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsysAdminTreeView.frx":2B38
            Key             =   "studyinprepwarning"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsysAdminTreeView.frx":2C92
            Key             =   "studyinprep"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsysAdminTreeView.frx":2DEC
            Key             =   "studyopenwarning"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsysAdminTreeView.frx":2F46
            Key             =   "studyopen"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsysAdminTreeView.frx":30A0
            Key             =   "studydistnoroles"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsysAdminTreeView.frx":31FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsysAdminTreeView.frx":3354
            Key             =   "userinactive"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsysAdminTreeView.frx":37A6
            Key             =   "useractive"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsysAdminTreeView.frx":3BF8
            Key             =   "databaseconnected"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView trvSysAdmin 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   12091
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imgIcons"
      Appearance      =   0
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "frmSysAdminTreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       frmSysAdminTreeView.frm
'   Author:     Richard Meinesz, October 2002
'   Purpose:    Tree view for System Managment.
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
'   Revisions:
'REM 12/03/04 - Added condition compilation arguments in DislpayPopupMenu routine so certain menu items are not displayed in the Desktop edition of MACRO
' MLM 20/06/05: bug 2500: tidied up trvSysAdmin_NodeClick
' TA 19/1/2006 - for performance password retries is no calculated out of the loop when building user list
' Mo 21/9/2007  Bug 2927, Overflow error when creating treeview of database studies & sites.
'               Loop variable in SiteWithStudies & StudyWithSites changed from Integer to Long.
'               modSysTreeViewSQL.StudySiteWithRoles has had a Distinct added to its SQL statement.
'               Note that the code in BuildStudiesSites needs to be rewritten, the section of code that builds the
'               Study/Site nodes for studies and sites with no roles assigned but have links in the TrialSite table
'               calls StudySiteWithRoles which will timeout when there are a large number of Studies & Sites.
'               BuildStudiesSites and the manner in which it works with StudySiteWithRoles need a rewrite.
'--------------------------------------------------------------------------------
Option Explicit

'constants for the node icon labels
Private Const msUSERS_ICON = "users"
Private Const msUSER_INACTIVE_ICON = "userinactive"
Private Const msUSER_ACTIVE_ICON = "useractive"
Private Const msUSER_ACTIVE_SYSADMIN_ICON = "usersys"
Private Const msSERVER_DBS_ICON = "databases"
Private Const msDB_CONNECTED_ICON = "databaseconnected"
Private Const msDB_LOGGED_IN_ICON = "databaseloggedin"
Private Const msDB_DISCONNECTED_ICON = "databasedisconnected"
Private Const msROLE_ICON = "role"
Private Const msROLES_ICON = "roles"
Private Const msSITES_ICON = "sites"
Private Const msACTIVE_SITE_ICON = "activesite"
Private Const msACTIVE_SITE_WARNING_ICON = "activesitewarning"
Private Const msINACTIVE_SITE_ICON = "inactivesite"
Private Const msSTUDIES_ICON = "studies"
Private Const msSTUDY_CLOSED_ICON = "studyclosed"
Private Const msSTUDY_CLOSED_WARNING_ICON = "studyclosedwarning"
Private Const msSTUDY_INPREP_WARNING_ICON = "studyinprepwarning"
Private Const msSTUDY_INPREP_ICON = "studyinprep"
Private Const msSTUDY_OPEN_WARNING_ICON = "studyopenwarning"
Private Const msSTUDY_OPEN_ICON = "studyopen"
Private Const msSTUDY_DIST_NOROLES_ICON = "studydistnoroles"

' constants for the node type tags in datalist
Private Const msUSERS_NODE = "Us"
Private Const msUSER_NODE = "U"
Private Const msSERVER_DATABASES_NODE = "Ds"
Private Const msDATABASE_NODE = "D"
Private Const msDISCONDATABASE_NODE = "DD"
Private Const msUPGRADE_DATABASE = "UG"
Private Const msSTUDIES_NODE = "Sts"
Private Const msSTUDY_NODE = "St"
Private Const msSITES_NODE = "Sis"
Private Const msSITE_NODE = "Si"
Private Const msROLES_NODE = "Rs"
Private Const msROLE_NODE = "R"

Private Const msSEPARATOR As String = "|"

'constants for the display text for each node
Private Const msSERVER_DBS_NODE_TEXT = "Server Databases"
Private Const msSERVER_DBS_NODE_LABEL = "Server DB Node"
Private Const msSTUDIES_NODE_TEXT = "Studies"
'Private Const msSTUDIES_NODE_LABEL = "Studies Node"
Private Const msSITES_NODE_TEXT = "Sites"
'Private Const msSITES_LABEL = "Sites Node"
Private Const msUSER_NODE_TEXT = "Users"
Private Const msUSER_NODE_LABEL = "Users Node"
Private Const msROLES_NODE_TEXT = "Roles"
Private Const msROLES_NODE_LABEL = "Roles Node"

'Key for the tree view nodes
Public Enum eIdType
    DatabaseCode = 0
    StudyId = 1
    StudyName = 2
    SiteCode = 3
    UserName = 4
    RoleCode = 5
End Enum

'key for node type
Public Enum eSMNodeTag
    UserTag = 0
    UsersTag = 1
    DatabaseTag = 2
    DatabasesTag = 3
    StudyTag = 4
    StudiesTag = 5
    SiteTag = 6
    SitesTag = 7
    RoleTag = 8
    RolesTag = 9
    UserRoleTag = 10
    DisconnectedDB = 11
    Upgrade = 12
End Enum

'status of distributed studies
Private Enum eSMStudyStatusDist
    ssdInPrep = 1
    ssdOpen = 2
    ssdClosedToRecruit = 3
    ssdClosedToFollowup = 4
    ssdSuspended = 5
    ssdDistributedNoRoles = 6
End Enum

'status of studies not distributed
Private Enum eSMStudyStatusNotDist
    ssndInPrep = 1
    ssndOpen = 2
    ssndClosedToRecruit = 3
    ssndClosedToFollowup = 4
    ssndSuspended = 5
End Enum


Private Enum eSMSiteStatus
    sisAcitive = 0
    sisInactive = 1
    sisActiveDistributedNoRoles = 2
    sisNotParticipatingInStudy = 3
    sisNotParticipatingInAnyStudies = 4
End Enum

Public Event Resize(sglWidth As Single)

Public Event SelectedNode(sNodeTag As eSMNodeTag, sDatabaseCode As String, lStudyId As Long, sStudyName As String, sSiteCode As String, sUsername As String, sRoleCode As String)

Private mconMACRO As ADODB.Connection
Private msglStartX As Single

Private mbViewUnusableRoles As Boolean

Private msDatabaseVersion As String

Private mcolDatabaseTags As Collection


'---------------------------------------------------------------------
Public Property Get ViewUnusableRoles() As Boolean
'---------------------------------------------------------------------

    ViewUnusableRoles = mbViewUnusableRoles

End Property

'---------------------------------------------------------------------
Public Property Let ViewUnusableRoles(bViewUnusableRoles As Boolean)
'---------------------------------------------------------------------
    
    If bViewUnusableRoles <> mbViewUnusableRoles Then
        Call RefreshTreeView
        mbViewUnusableRoles = bViewUnusableRoles
    End If

End Property

'---------------------------------------------------------------------
Public Property Get DatabaseTags() As Collection
'---------------------------------------------------------------------

    Set DatabaseTags = mcolDatabaseTags

End Property

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    Me.Icon = frmMenu.Icon
    
    ' initialise image list
    trvSysAdmin.ImageList = imgIcons
    
    mbViewUnusableRoles = False 'Need to get this from the settings file???

    Call RefreshTreeView
    
    trvSysAdmin.DragMode = 0
    
    cmdSep.Top = 0
    trvSysAdmin.Top = 0
    trvSysAdmin.Left = 0
    
    msglStartX = -1 '=not dragging
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "Form_Load")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
End Sub


'---------------------------------------------------------------------
Private Sub Form_Resize()
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    cmdSep.Left = Me.Width - cmdSep.Width
    trvSysAdmin.Width = cmdSep.Left
    cmdSep.Height = Me.Height
    trvSysAdmin.Height = Me.Height
    
End Sub


'---------------------------------------------------------------------
Public Sub RefreshTreeView(Optional bExpand As Boolean = False)
'---------------------------------------------------------------------
'REM 07/10/02
'Refresh the tree view
' TA if bexpand is set to false the node for each db is not built
'---------------------------------------------------------------------

    HourglassOn
    
    If Not bExpand Then
        'check setttings file whether to expand
        bExpand = (GetMACROSetting("SMBuildTree", "true") <> "false")
    End If
    
    Call LockWindow(trvSysAdmin)
    
    Call BuildDatabasesNodes(bExpand)
    'Call BuildDatabaseStudySite
    
    Call BuildUserNodes
    
    Call BuildRoleNodes
    
    Call UnlockWindow
    
    HourglassOff

End Sub

'---------------------------------------------------------------------
Private Sub BuildDatabasesNodes(bExpand As Boolean)
'---------------------------------------------------------------------
'REM 07/10/03
'Build the Server database nodes
'---------------------------------------------------------------------
Dim nodx As Node
Dim vDatabases As Variant
Dim i As Integer
Dim oDatabase As MACROUserBS30.Database
Dim conMACROADODBConnection As ADODB.Connection
Dim sConnectionString As String
Dim sDatabaseCode As String
Dim sMessage As String
Dim sToolTip As String
Dim sImage As String
Dim bProceedWithConnection As Boolean
Dim sDBTypeAndVersion As String

    On Error GoTo ErrLabel
    
    trvSysAdmin.Nodes.Clear

    'Create the ServerDatabases node
    sToolTip = msSERVER_DBS_NODE_LABEL
    Set nodx = trvSysAdmin.Nodes.Add(, tvwLast, GetDatabaseNodeKey(""), msSERVER_DBS_NODE_TEXT, msSERVER_DBS_ICON)
        nodx.Tag = msSERVER_DATABASES_NODE & msSEPARATOR & sToolTip
    

        Set mcolDatabaseTags = New Collection
        
        'get all the databases in the Security database
        vDatabases = Databases
        If Not IsNull(vDatabases) Then
            'loop through all databases building study site nodes
            For i = 0 To UBound(vDatabases, 2)
                
                sDatabaseCode = vDatabases(0, i)
                If Not bExpand Then
                    'couldn't find server
                    sToolTip = sDatabaseCode
                    sImage = msDB_DISCONNECTED_ICON
                    'add the database node
                    Set nodx = trvSysAdmin.Nodes.Add(GetDatabaseNodeKey(""), tvwChild, _
                    GetDatabaseNodeKey(sDatabaseCode), sDatabaseCode, sImage)
                    nodx.Tag = msDISCONDATABASE_NODE & msSEPARATOR & sToolTip
                    
                Else
                    sDBTypeAndVersion = CheckDBTypeAndVersion(sDatabaseCode)
                    If sDBTypeAndVersion = "" Then 'then either database cannot be connected to or it is a current database
                        Set oDatabase = New MACROUserBS30.Database
                        'load up database properties
                        If oDatabase.Load(SecurityADODBConnection, goUser.UserName, sDatabaseCode, "", False, sMessage) Then
                        
                            sConnectionString = oDatabase.ConnectionString
                        
                            'attempt to create connection
                            'if connection is created then proceed
                            If ConnectionCreated(sConnectionString, sDatabaseCode, sMessage) Then
                            
                            
                                If sDatabaseCode = goUser.Database.DatabaseCode Then
                                    sToolTip = "Logged into " & sDatabaseCode & " database"
                                    sImage = msDB_LOGGED_IN_ICON
                                Else
                                    sToolTip = sDatabaseCode & " is connected"
                                    sImage = msDB_CONNECTED_ICON
                                End If
                                
                                'add the database node
                                Set nodx = trvSysAdmin.Nodes.Add(GetDatabaseNodeKey(""), tvwChild, _
                                GetDatabaseNodeKey(sDatabaseCode), sDatabaseCode, sImage)
                                nodx.Tag = msDATABASE_NODE & msSEPARATOR & sToolTip
                                
                                'builds the study/site node for each database
                                Call BuildStudySiteNode(sDatabaseCode)
                            Else
                                sToolTip = sMessage
                                sImage = msDB_DISCONNECTED_ICON
                                'add the database node
                                Set nodx = trvSysAdmin.Nodes.Add(GetDatabaseNodeKey(""), tvwChild, _
                                GetDatabaseNodeKey(sDatabaseCode), sDatabaseCode, sImage)
                                nodx.Tag = msDISCONDATABASE_NODE & msSEPARATOR & sToolTip
                            End If
                            
                        Else 'if connection fails then add a disconnected DB icon and tool tip
                            sToolTip = sMessage
                            sImage = msDB_DISCONNECTED_ICON
                            'add the database node
                            Set nodx = trvSysAdmin.Nodes.Add(GetDatabaseNodeKey(""), tvwChild, _
                            GetDatabaseNodeKey(sDatabaseCode), sDatabaseCode, sImage)
                            nodx.Tag = msDISCONDATABASE_NODE & msSEPARATOR & sToolTip
                        End If
                    Else 'old MACRO version so give it the upgrade tag and tooltip
                        sToolTip = sDBTypeAndVersion
                        sImage = msDB_DISCONNECTED_ICON
                        'add the database node
                        Set nodx = trvSysAdmin.Nodes.Add(GetDatabaseNodeKey(""), tvwChild, _
                        GetDatabaseNodeKey(sDatabaseCode), sDatabaseCode, sImage)
                        nodx.Tag = msUPGRADE_DATABASE & msSEPARATOR & sToolTip
                    End If
                End If
                
                mcolDatabaseTags.Add nodx.Tag, sDatabaseCode
                
            Next
    
            'make sure the Server Database node is expanded
            nodx.EnsureVisible
        End If
    
    
Exit Sub
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "frmSysAdminTreeView.BuildDatabaseStudySite"
End Sub

'---------------------------------------------------------------------
Private Sub BuildStudySiteNode(sDatabaseCode As String)
'---------------------------------------------------------------------
'REM 07/10/03
'Build the Study and site node
'---------------------------------------------------------------------
Dim nodx As Node
Dim sToolTip As String

    On Error GoTo ErrLabel

    'Create the Studies node
    sToolTip = "All studies in " & sDatabaseCode
    Set nodx = trvSysAdmin.Nodes.Add(GetDatabaseNodeKey(sDatabaseCode), tvwChild, _
                GetStudyNodeKey(msSTUDIES_NODE, sDatabaseCode, -1, ""), msSTUDIES_NODE_TEXT, msSTUDIES_ICON)
    nodx.Tag = msSTUDIES_NODE & msSEPARATOR & sToolTip
    
    'Create the Sites node
    sToolTip = "All sites in " & sDatabaseCode
    Set nodx = trvSysAdmin.Nodes.Add(GetDatabaseNodeKey(sDatabaseCode), tvwChild, _
                GetSiteNodeKey(msSITES_NODE, sDatabaseCode, ""), msSITES_NODE_TEXT, msSITES_ICON)
    nodx.Tag = msSITES_NODE & msSEPARATOR & sToolTip
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmSysAdminTreeView.BuildStudySiteNode"
End Sub

'---------------------------------------------------------------------
Private Sub BuildDatabaseStudiesSites(sDatabaseCode As String)
'---------------------------------------------------------------------
'REM 07/10/03
'Build the Study and site nodes when user clicks on a database
'---------------------------------------------------------------------
Dim nDatabaseType As Integer
Dim sErrMsg As String

    HourglassOn
    
    Call SetConnection(SecurityADODBConnection, sDatabaseCode, nDatabaseType, sErrMsg)
    Call BuildStudiesSites(sDatabaseCode, nDatabaseType)
    
    HourglassOff
    
End Sub
    
''---------------------------------------------------------------------
'Private Sub BuildDatabaseStudySite()
''---------------------------------------------------------------------
''REM 03/10/02
''Build the Server database node with all the database nodes
''---------------------------------------------------------------------
'Dim nodx As Node
'Dim vDatabases As Variant
'Dim i As Integer
'Dim oDatabase As MACROUserBS30.Database
'Dim conMACROADODBConnection As ADODB.Connection
'Dim sConnectionString As String
'Dim sDatabaseCode As String
'Dim sMessage As String
'Dim sToolTip As String
'Dim sImage As String
'Dim bProceedWithConnection As Boolean
'Dim sDBTypeAndVersion As String
'
'    On Error GoTo ErrLabel
'
'    trvSysAdmin.Nodes.Clear
'
'    'Create the ServerDatabases node
'    sToolTip = msSERVER_DBS_NODE_LABEL
'    Set nodx = trvSysAdmin.Nodes.Add(, tvwLast, GetDatabaseNodeKey(""), msSERVER_DBS_NODE_TEXT, msSERVER_DBS_ICON)
'        nodx.Tag = msSERVER_DATABASES_NODE & msSEPARATOR & sToolTip
'
'    Set mcolDatabaseTags = New Collection
'
'    'get all the databases in the Security database
'    vDatabases = Databases
'    If Not IsNull(vDatabases) Then
'        'loop through all databases building study site nodes
'        For i = 0 To UBound(vDatabases, 2)
'
'            sDatabaseCode = vDatabases(0, i)
'
'            'TA 18/12/2002: check whether server exists/is on network
'            If vDatabases(2, i) = MACRODatabaseType.sqlserver Then
'                bProceedWithConnection = True '(ServerExists(CStr(vDatabases(1, i)), "Shell.Application") <> serNo)
'            Else
'                bProceedWithConnection = True
'            End If
'
'            If Not bProceedWithConnection Then
'                'couldn't find server
'                sToolTip = "Cannot find server " & DOUBLE_QUOTE & CStr(vDatabases(1, i)) & DOUBLE_QUOTE
'                sImage = msDB_DISCONNECTED_ICON
'                'add the database node
'                Set nodx = trvSysAdmin.Nodes.Add(GetDatabaseNodeKey(""), tvwChild, _
'                GetDatabaseNodeKey(sDatabaseCode), sDatabaseCode, sImage)
'                nodx.Tag = msDISCONDATABASE_NODE & msSEPARATOR & sToolTip
'
'            Else
'                sDBTypeAndVersion = CheckDBTypeAndVersion(sDatabaseCode)
'                If sDBTypeAndVersion = "" Then 'then either database cannot be connected to or it is a current database
'                    Set oDatabase = New MACROUserBS30.Database
'                    'load up database properties
'                    If oDatabase.Load(SecurityADODBConnection, goUser.UserName, sDatabaseCode, "", False, sMessage) Then
'
'                        sConnectionString = oDatabase.ConnectionString
'
'                        'attempt to create connection
'                        'if connection is created then proceed
'                        If ConnectionCreated(sConnectionString, sDatabaseCode, sMessage) Then
'
'
'                            If sDatabaseCode = goUser.Database.DatabaseCode Then
'                                sToolTip = "Logged into " & sDatabaseCode & " database"
'                                sImage = msDB_LOGGED_IN_ICON
'                            Else
'                                sToolTip = sDatabaseCode & " is connected"
'                                sImage = msDB_CONNECTED_ICON
'                            End If
'
'                            'add the database node
'                            Set nodx = trvSysAdmin.Nodes.Add(GetDatabaseNodeKey(""), tvwChild, _
'                            GetDatabaseNodeKey(sDatabaseCode), sDatabaseCode, sImage)
'                            nodx.Tag = msDATABASE_NODE & msSEPARATOR & sToolTip
'
'                            'builds the study and site nodes
'                            Call BuildStudiesSites(sDatabaseCode)
'                        Else
'                            sToolTip = sMessage
'                            sImage = msDB_DISCONNECTED_ICON
'                            'add the database node
'                            Set nodx = trvSysAdmin.Nodes.Add(GetDatabaseNodeKey(""), tvwChild, _
'                            GetDatabaseNodeKey(sDatabaseCode), sDatabaseCode, sImage)
'                            nodx.Tag = msDISCONDATABASE_NODE & msSEPARATOR & sToolTip
'                        End If
'
'                    Else 'if connection fails then add a disconnected DB icon and tool tip
'                        sToolTip = sMessage
'                        sImage = msDB_DISCONNECTED_ICON
'                        'add the database node
'                        Set nodx = trvSysAdmin.Nodes.Add(GetDatabaseNodeKey(""), tvwChild, _
'                        GetDatabaseNodeKey(sDatabaseCode), sDatabaseCode, sImage)
'                        nodx.Tag = msDISCONDATABASE_NODE & msSEPARATOR & sToolTip
'                    End If
'                Else 'old MACRO version so give it the upgrade tag and tooltip
'                    sToolTip = sDBTypeAndVersion
'                    sImage = msDB_DISCONNECTED_ICON
'                    'add the database node
'                    Set nodx = trvSysAdmin.Nodes.Add(GetDatabaseNodeKey(""), tvwChild, _
'                    GetDatabaseNodeKey(sDatabaseCode), sDatabaseCode, sImage)
'                    nodx.Tag = msUPGRADE_DATABASE & msSEPARATOR & sToolTip
'                End If
'            End If
'
'            mcolDatabaseTags.Add nodx.Tag, sDatabaseCode
'
'        Next
'
'        'make sure the Server Database node is expanded
'        nodx.EnsureVisible
'    End If
'
'
'Exit Sub
'ErrLabel:
'Err.Raise Err.Number, , Err.Description & "|" & "frmSysAdminTreeView.BuildDatabaseStudySite"
'End Sub

'---------------------------------------------------------------------
Private Function CheckDBTypeAndVersion(sDatabaseCode As String) As String
'---------------------------------------------------------------------
'REM 18/06/03
'Routine to check database type and version
'---------------------------------------------------------------------
Dim sVersion As String
Dim sSubVersion As String
Dim bAccess As Boolean

    Call DatabaseVersionSubVersion(sDatabaseCode, sVersion, sSubVersion, bAccess)
    
    If (sVersion <> goUser.SecurityDBVersion) Or (sSubVersion <> goUser.SecurityDBSubVersion) Then
        If bAccess Then
            CheckDBTypeAndVersion = sDatabaseCode & " is an Access database and must be upgraded to a MACRO " & goUser.SecurityDBVersion & "." & goUser.SecurityDBSubVersion & " database."

        Else
        
            CheckDBTypeAndVersion = sDatabaseCode & " database version is " & sVersion & "." & sSubVersion & " and must be upgraded to " & goUser.SecurityDBVersion & "." & goUser.SecurityDBSubVersion

        End If
    Else
    
        CheckDBTypeAndVersion = ""

    End If

End Function

'---------------------------------------------------------------------
Private Function ConnectionCreated(sConnectionString As String, sDatabaseCode As String, ByRef sMessage As String) As Boolean
'---------------------------------------------------------------------
'REM 29/10/02
'Creates modular level connection to a MACRO database, if connection fails for any reason will return false
'and the error message
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsVersion As ADODB.Recordset
Dim sVersion As String
Dim sSubVersion As String

    On Error GoTo ErrLabel
    
    'create connection for selected database
    Set mconMACRO = New ADODB.Connection
    
    mconMACRO.Open sConnectionString
    mconMACRO.CursorLocation = adUseClient
    
    ConnectionCreated = True
'    sSQL = "SELECT * FROM MACROControl"
'    Set rsVersion = New ADODB.Recordset
'    rsVersion.Open sSQL, mconMACRO, adOpenKeyset, adLockReadOnly, adCmdText
'
'    sVersion = rsVersion![MACROVersion]
'    sSubVersion = rsVersion![BuildSubVersion]

'    Call DatabaseVersionSubVersion(sDatabaseCode, sVersion, sSubVersion)
'
'    If (sVersion <> goUser.SecurityDBVersion) Or (sSubVersion <> goUser.SecurityDBSubVersion) Then
'        If sVersion = msACCESS_DB Then
'            sMessage = sDatabaseCode & " database version is an Access database and must be upgraded to " & goUser.SecurityDBVersion & "." & goUser.SecurityDBSubVersion
'            ConnectionCreated = False
'        Else
'
'            sMessage = sDatabaseCode & " database version is " & sVersion & "." & sSubVersion & " and must be upgraded to " & goUser.SecurityDBVersion & "." & goUser.SecurityDBSubVersion
'            ConnectionCreated = False
'        End If
'    Else
'
'        sMessage = ""
'        ConnectionCreated = True
'    End If
    
Exit Function
ErrLabel:
    'if the connection fails then return the error message
    sMessage = "Could not create connection to " & sDatabaseCode & " because " & Err.Description
    ConnectionCreated = False
End Function

'----------------------------------------------------------------------------------------'
Private Function SetConnection(oSecCon As ADODB.Connection, sDatabaseCode As String, _
                               ByRef nDatabaseType As Integer, ByRef sErrorMsg As String) As String
'----------------------------------------------------------------------------------------'
'REM 11/11/02
'Create a connection for a given database
'----------------------------------------------------------------------------------------'
Dim conMACRO As ADODB.Connection
Dim sConnection As String

    On Error GoTo ErrLabel
    
    sConnection = ConnectionString(sDatabaseCode, nDatabaseType, sErrorMsg, oSecCon)
        
    If sConnection <> "" Then
        
        If mconMACRO.State = adStateOpen Then
            mconMACRO.Close
        End If
        
        'create connection for selected database
        Set conMACRO = New ADODB.Connection
    
        mconMACRO.Open sConnection
        mconMACRO.CursorLocation = adUseClient
    
    End If
    
    SetConnection = sConnection
    
Exit Function
ErrLabel:
    sErrorMsg = "MACRO Database connection error: " & Err.Description & ": Error no. " & Err.Number
End Function

'---------------------------------------------------------------------
Private Function ConnectionString(sDatabaseCode As String, ByRef nDatabaseType As Integer, ByRef sErrorMsg As String, Optional oSecCon As ADODB.Connection = Nothing) As String
'---------------------------------------------------------------------
'REM 26/11/02
'Create a MACRO database connection string for a given database
'---------------------------------------------------------------------
Dim oDatabase As MACROUserBS30.Database
Dim sConnection As String

    On Error GoTo ErrLabel

    Set oDatabase = New MACROUserBS30.Database

    If oDatabase.Load(oSecCon, "", sDatabaseCode, "", False, sErrorMsg) Then
        nDatabaseType = oDatabase.DatabaseType
        ConnectionString = oDatabase.ConnectionString
    Else
        nDatabaseType = -1
        ConnectionString = ""
    End If
    
    Set oDatabase = Nothing

Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "basSystemTransfer.ConnectionString"
End Function

'---------------------------------------------------------------------
Private Sub DatabaseVersionSubVersion(sDatabaseCode As String, ByRef sVersion As String, ByRef sSubVersion As String, ByRef bAccess As Boolean)
'---------------------------------------------------------------------
'REM 18/06/03
'Returns a a databases version and sub version
'---------------------------------------------------------------------
Dim oDatabase As MACROUserBS30.Database
Dim rsVersion As ADODB.Recordset
Dim rsDBType As ADODB.Recordset
Dim conMACRO As ADODB.Connection
Dim sSQL As String
Dim sCon As String
Dim sMessage As String
Dim nDatabaseType As Integer
Dim lErrNo As Long
    
    On Error GoTo ErrLabel
    
    sSQL = "SELECT DatabaseType FROM Databases WHERE DatabaseCode = '" & sDatabaseCode & "'"
    Set rsDBType = New ADODB.Recordset
    rsDBType.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    nDatabaseType = rsDBType!DatabaseType
    
    rsDBType.Close
    Set rsDBType = Nothing
    
    Select Case nDatabaseType
    Case MACRODatabaseType.Access
        bAccess = True
        sVersion = ""
        sSubVersion = ""
    Case Else
        bAccess = False
        Set oDatabase = New MACROUserBS30.Database
        
        Call oDatabase.Load(SecurityADODBConnection, goUser.UserName, sDatabaseCode, "", False, sMessage)
        sCon = oDatabase.ConnectionString
        On Error Resume Next 'resume after ERROR
        Set conMACRO = New ADODB.Connection
        conMACRO.Open sCon
        conMACRO.CursorLocation = adUseClient
        lErrNo = Err.Number
        Err.Clear
        
        On Error GoTo ErrLabel
        
        If lErrNo <> 0 Then 'if there is a connection error we set the version to current and catch the error later
            sVersion = goUser.SecurityDBVersion
            sSubVersion = goUser.SecurityDBSubVersion
        Else
            sSQL = "SELECT * FROM MACROControl"
            Set rsVersion = New ADODB.Recordset
            rsVersion.Open sSQL, conMACRO, adOpenKeyset, adLockReadOnly, adCmdText
            
            sVersion = rsVersion![MACROVersion]
            sSubVersion = rsVersion![BuildSubVersion]
            
            Set rsVersion = Nothing
            conMACRO.Close
            Set conMACRO = Nothing
        End If
    

        'rsVersion.Close

    End Select
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmSysAdminTreeView.DatabaseVersionSubVersion"
End Sub

'---------------------------------------------------------------------
Private Sub BuildStudiesSites(sDatabaseCode As String, nDatabaseType As Integer)
'---------------------------------------------------------------------
'REM 14/10/02
'Build the Study and site nodes
'---------------------------------------------------------------------
Dim nodx As Node
Dim vStudies As Variant
Dim vSites As Variant
Dim vStudySiteWithRoles As Variant
Dim vStudySiteWithoutRoles As Variant
Dim vStudySiteNotDist As Variant
Dim sStudyStatus As String
Dim sSiteStudyStatus As String
Dim sStudySiteStatus As String
Dim sSiteStatus
Dim sToolTip As String

    On Error GoTo ErrLabel

    'Create the Studies node
    sToolTip = "All studies in " & sDatabaseCode
    Set nodx = trvSysAdmin.Nodes.Add(GetDatabaseNodeKey(sDatabaseCode), tvwChild, _
                GetStudyNodeKey(msSTUDIES_NODE, sDatabaseCode, -1, ""), msSTUDIES_NODE_TEXT, msSTUDIES_ICON)
    nodx.Tag = msSTUDIES_NODE & msSEPARATOR & sToolTip
    
    'Create the Sites node
    sToolTip = "All sites in " & sDatabaseCode
    Set nodx = trvSysAdmin.Nodes.Add(GetDatabaseNodeKey(sDatabaseCode), tvwChild, _
                GetSiteNodeKey(msSITES_NODE, sDatabaseCode, ""), msSITES_NODE_TEXT, msSITES_ICON)
    nodx.Tag = msSITES_NODE & msSEPARATOR & sToolTip
    
    'Build the Study/Site nodes for studies and sites with roles assigned
    vStudySiteWithRoles = StudySiteWithRoles(mconMACRO)
    If Not IsNull(vStudySiteWithRoles) Then
        Call StudyWithSites(vStudySiteWithRoles, sDatabaseCode)
        Call SiteWithStudies(vStudySiteWithRoles, sDatabaseCode)
    End If
    
    'Mo 21/9/2007  Bug 2927, the following section of code could do with being rewritten, it calls StudySiteWithoutRoles,
    'which can timeout when there are a large number of Studies and sites.
    'Build the Study/Site nodes for studies and sites with NO roles assigned but have links in the TrialSite table
    vStudySiteWithoutRoles = StudySiteWithoutRoles(mconMACRO, nDatabaseType)
    If Not IsNull(vStudySiteWithoutRoles) Then
        sStudyStatus = eSMStudyStatusDist.ssdDistributedNoRoles 'any status, distributed but no roles
        sStudySiteStatus = 2 ' site with no user roles
        Call StudyWithSites(vStudySiteWithoutRoles, sDatabaseCode, "", sStudySiteStatus)
        Call SiteWithStudies(vStudySiteWithoutRoles, sDatabaseCode, sSiteStudyStatus, sStudyStatus)
    End If
    
    'if true the tree is built with study/site combinations that have been assigned user roles but
    ' not linked in the TrialSite table
    If mbViewUnusableRoles = True Then
    
        'Build the Study/Site nodes for study-site combinations that have roles but are not distributed
        vStudySiteNotDist = StudySiteNotDist(mconMACRO, nDatabaseType)
        If Not IsNull(vStudySiteNotDist) Then
            sStudyStatus = ""
            sStudySiteStatus = 3 'site with roles assigned but not linked to a study
            Call StudyWithSites(vStudySiteNotDist, sDatabaseCode, sStudyStatus, sStudySiteStatus, True)
            Call SiteWithStudies(vStudySiteNotDist, sDatabaseCode, sSiteStudyStatus, sStudyStatus, True)
        End If
    End If
    
    'Build the Study nodes for studies that only occur in the ClinicalTrial table table
    'i.e. no roles assigned or links in the trailsite table
    'returns all the studies but StudyWithSites routine only adds nodes not already there
    vStudies = Studies(mconMACRO)
    If Not IsNull(vStudies) Then
        Call StudyWithSites(vStudies, sDatabaseCode, "", "", True, True)
    End If
    'Build the Site nodes for sites that only occur in the site table
    'i.e. no roles assigned or links in the trailsite table
    'returns all the sites but SiteWithStudies routine only adds nodes not already there
    vSites = Sites(mconMACRO)
    If Not IsNull(vSites) Then
        sSiteStatus = eSMSiteStatus.sisNotParticipatingInAnyStudies
        Call SiteWithStudies(vSites, sDatabaseCode, sSiteStatus, "", True, True)
    End If
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmSysAdminTreeView.BuildStudiesSites"
End Sub

'---------------------------------------------------------------------
Private Sub StudyWithSites(vStudyWithSites As Variant, sDatabaseCode As String, Optional sStatusStudy As String = "", _
                           Optional sStatusSite As String = "", Optional bStudyNotDistributed As Boolean = False, _
                           Optional bStudyNoSites As Boolean = False)
'---------------------------------------------------------------------
'REM 15/10/02
'Builds Studies with accociated sites
'---------------------------------------------------------------------
Dim nodx As Node
'Mo 21/9/2007 Bug 2927
Dim i As Long
Dim sStudyStatusIcon As String
Dim sStudyName As String
Dim lStudyId As Long
Dim sSiteCode As String
Dim sSiteStatus As String
Dim sStudyStatus As String
Dim sToolTip As String

    On Error GoTo ErrLabel

    'build each Study node with its accociated Site nodes
    For i = 0 To UBound(vStudyWithSites, 2)
            
        sStudyName = vStudyWithSites(1, i)
        lStudyId = vStudyWithSites(0, i)
        sSiteCode = vStudyWithSites(3, i)
    
        'SET STUDY STATUS
        'if bStudySiteNotDist is true then the study has not been distributed and the status has to be set accordingly
        If bStudyNotDistributed = True Then
            sStudyStatus = vStudyWithSites(2, i)
            Select Case sStudyStatus
            Case eSMStudyStatusNotDist.ssndInPrep 'Is = 1 'Status of in Prep with roles, not distributed
                sStudyStatusIcon = msSTUDY_INPREP_WARNING_ICON
            Case eSMStudyStatusNotDist.ssndOpen 'Is = 2 'Status of open, with roles, not distributed
                sStudyStatusIcon = msSTUDY_OPEN_WARNING_ICON
            Case eSMStudyStatusNotDist.ssndClosedToFollowup, eSMStudyStatusNotDist.ssndClosedToRecruit, eSMStudyStatusNotDist.ssndSuspended 'Is = 3 Or 4 'status closed with roles, not distributed
                sStudyStatusIcon = msSTUDY_CLOSED_WARNING_ICON
            End Select
            sToolTip = "No sites are participating in Study " & sStudyName
        Else

            If sStatusStudy = "" Then
                sStudyStatus = vStudyWithSites(2, i)
            Else
                sStudyStatus = sStatusStudy
            End If
        
            'Status 1 to 5 get from DB, others set by being passed in
            Select Case sStudyStatus
            Case eSMStudyStatusDist.ssdInPrep ' Is = 1 ' study in prep
                sStudyStatusIcon = msSTUDY_INPREP_ICON
                sToolTip = "Study " & sStudyName & " is in preparation"
            Case eSMStudyStatusDist.ssdOpen 'Is = 2 'study open
                sStudyStatusIcon = msSTUDY_OPEN_ICON
                sToolTip = "Study " & sStudyName & " is open"
            Case eSMStudyStatusDist.ssdClosedToRecruit 'Is = 3 'study closed to recruitment
                sStudyStatusIcon = msSTUDY_CLOSED_ICON
                sToolTip = "Study " & sStudyName & " is closed to recruitment"
            Case eSMStudyStatusDist.ssdClosedToFollowup 'Is = 4 'closed to follow up
                sStudyStatusIcon = msSTUDY_CLOSED_ICON
                sToolTip = "Study " & sStudyName & " is closed to follow up"
            Case eSMStudyStatusDist.ssdSuspended '5 suspended
                sStudyStatusIcon = msSTUDY_CLOSED_ICON
                sToolTip = "Study " & sStudyName & " is suspended"
            Case eSMStudyStatusDist.ssdDistributedNoRoles 'Is = 6 'distribute with no roles
                sStudyStatusIcon = msSTUDY_DIST_NOROLES_ICON
                sToolTip = sStudyName & " - " & sSiteCode & " combination has no user roles assigned"
            End Select
        End If
        
        If sStatusSite = "" Then
            sSiteStatus = vStudyWithSites(4, i)
        Else
            sSiteStatus = sStatusSite
        End If
        
        'if study does not exist then add the node
        If Not DoesNodeExist(GetStudyNodeKey(msSTUDIES_NODE, sDatabaseCode, lStudyId, sStudyName)) Then
        
            Set nodx = trvSysAdmin.Nodes.Add(GetStudyNodeKey(msSTUDIES_NODE, sDatabaseCode, -1, ""), tvwChild, _
                GetStudyNodeKey(msSTUDIES_NODE, sDatabaseCode, lStudyId, sStudyName), sStudyName, sStudyStatusIcon)
            nodx.Tag = msSTUDY_NODE & msSEPARATOR & sToolTip
            
        End If
        
        If Not bStudyNoSites Then
            'add the site nodes to the study
            Call BuildSiteNodesForStudy(GetStudyNodeKey(msSTUDIES_NODE, sDatabaseCode, lStudyId, sStudyName), _
                                        GetSiteNodeKey(msSITES_NODE, sDatabaseCode, sSiteCode, lStudyId, sStudyName), _
                                        sSiteCode, sStudyName, sSiteStatus)
        End If
        
    Next

Exit Sub
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "frmSysAdminTreeView.StudyWithSites"
End Sub

'---------------------------------------------------------------------
Private Sub SiteWithStudies(vStudyWithSites As Variant, sDatabaseCode As String, Optional ByVal sStatusSite As String = "", _
                            Optional sStatusStudy As String = "", Optional ByVal bStudyNotDistributed As Boolean = False, _
                            Optional bSiteNoStudies As Boolean = False)
'---------------------------------------------------------------------
'REM 15/10/02
'build the site nodes and add the studies for each site
'---------------------------------------------------------------------
Dim nodx As Node
'Mo 21/9/2007 Bug 2927
Dim i As Long
Dim lStudyId As Long
Dim sSiteCode As String
Dim sSiteStatusIcon As String
Dim sStudyName As String
Dim sStudyStatus As String
Dim sSiteStatus As String
Dim sToolTip As String
    
    On Error GoTo ErrLabel

    'build the site nodes and add the studies for each site
    For i = 0 To UBound(vStudyWithSites, 2)
    
        sSiteCode = vStudyWithSites(3, i)
        sStudyName = vStudyWithSites(1, i)
        lStudyId = vStudyWithSites(0, i)
        
        'Check Site Stautus and set Icon
        If sStatusSite = "" Then
            sSiteStatus = vStudyWithSites(4, i)
        Else
            sSiteStatus = sStatusSite
        End If
        
        Select Case sSiteStatus
        Case eSMSiteStatus.sisAcitive '0  'active site
            sSiteStatusIcon = msACTIVE_SITE_ICON
            sToolTip = "Site " & sSiteCode & " is active"
        Case eSMSiteStatus.sisInactive '1
            'Inactive site
            sSiteStatusIcon = msINACTIVE_SITE_ICON
            sToolTip = "Site " & sSiteCode & " is inactive"
        Case eSMSiteStatus.sisActiveDistributedNoRoles '2
            'active site distributed but no user roles
            sSiteStatusIcon = msACTIVE_SITE_WARNING_ICON
            sToolTip = sStudyName & " - " & sSiteCode & " combination has no user roles assigned"
        Case eSMSiteStatus.sisNotParticipatingInStudy '3
            'Site with user roles assigned but not participating in a specific study
            sSiteStatusIcon = msINACTIVE_SITE_ICON
            sToolTip = "Site " & sSiteCode & " is not participating study " & sStudyName
        Case eSMSiteStatus.sisNotParticipatingInAnyStudies '4
            'A site that is not participating in any studies
            sSiteStatusIcon = msINACTIVE_SITE_ICON
            sToolTip = "Site " & sSiteCode & " is not participating in any studies"
        End Select
        
        If sStatusStudy = "" Then
            sStudyStatus = vStudyWithSites(2, i)
        Else
            sStudyStatus = sStatusStudy
        End If
        
        If Not DoesNodeExist(GetSiteNodeKey(msSITES_NODE, sDatabaseCode, sSiteCode)) Then
            'add each site node
            Set nodx = trvSysAdmin.Nodes.Add(GetSiteNodeKey(msSITES_NODE, sDatabaseCode, ""), tvwChild, _
                        GetSiteNodeKey(msSITES_NODE, sDatabaseCode, sSiteCode), sSiteCode, sSiteStatusIcon)
            nodx.Tag = msSITES_NODE & msSEPARATOR & sToolTip
        End If
        
        If Not bSiteNoStudies Then
            'add the study nodes to each site
            Call BuildStudyNodesForSite(GetSiteNodeKey(msSITES_NODE, sDatabaseCode, sSiteCode), GetStudyNodeKey(msSTUDIES_NODE, sDatabaseCode, lStudyId, _
                                    sStudyName, sSiteCode), sStudyStatus, sStudyName, sSiteCode, bStudyNotDistributed)
        End If
    Next

Exit Sub
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "frmSysAdminTreeView.SiteWithStudies"
End Sub

'---------------------------------------------------------------------
Private Function DoesNodeExist(Key As String) As Boolean
'---------------------------------------------------------------------
'REM 15/10/02
'Checks for the existance of a node
'---------------------------------------------------------------------
Dim oNode As Node

    On Error Resume Next
    Set oNode = trvSysAdmin.Nodes(Key)
    Select Case Err.Number
    Case 0
        DoesNodeExist = True
    Case 35601
        DoesNodeExist = False
    Case Else
        Err.Raise Err.Number
    End Select

End Function

'---------------------------------------------------------------------
Private Sub BuildSiteNodesForStudy(sStudyNodeKey As String, sSiteNodeKey As String, sSiteCode As String, sStudyName As String, _
                                   sSiteStatus As String)
'---------------------------------------------------------------------
'REM 03/10/02
'Build the Site nodes for a specific study
'---------------------------------------------------------------------
Dim nodx As Node
Dim sSiteStatusIcon As String
Dim sToolTip As String
    
    On Error GoTo ErrLabel

        Select Case sSiteStatus
        Case eSMSiteStatus.sisAcitive '0  'active site
            sSiteStatusIcon = msACTIVE_SITE_ICON
            sToolTip = "Site " & sSiteCode & " is active"
        Case eSMSiteStatus.sisInactive '1
            'Inactive site
            sSiteStatusIcon = msINACTIVE_SITE_ICON
            sToolTip = "Site " & sSiteCode & " is inactive"
        Case eSMSiteStatus.sisActiveDistributedNoRoles '2
            'active site distributed but no user roles
            sSiteStatusIcon = msACTIVE_SITE_WARNING_ICON
            sToolTip = sStudyName & " - " & sSiteCode & " combination has no user roles assigned"
        Case eSMSiteStatus.sisNotParticipatingInStudy '3
            'Site with user roles assigned but not participating in a study
            sSiteStatusIcon = msINACTIVE_SITE_ICON
            sToolTip = "Site " & sSiteCode & " is not participating in study " & sStudyName
        End Select
    
    If Not DoesNodeExist(sSiteNodeKey) Then
        'add the site node
        Set nodx = trvSysAdmin.Nodes.Add(sStudyNodeKey, tvwChild, _
            sSiteNodeKey, sSiteCode, sSiteStatusIcon)
        nodx.Tag = msSITE_NODE & msSEPARATOR & sToolTip
    End If
        
Exit Sub
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "frmSysAdminTreeView.BuildStudySitesNodes"
End Sub

'---------------------------------------------------------------------
Private Sub BuildStudyNodesForSite(sSiteNodeKey As String, sStudyNodeKey As String, sStudyStatus As String, _
                                   sStudyName As String, sSiteCode As String, bStudyNotDistributed As Boolean)
'---------------------------------------------------------------------
'REM 07/10/02
'Build the study nodes for a specific site
'---------------------------------------------------------------------
Dim nodx As Node
Dim sStudyStatusIcon As String
Dim sToolTip As String


    On Error GoTo ErrLabel
        
        'SET STUDY STATUS
        'if bStudySiteNotDist is true then the study has not been distributed and the status has to be set accordingly
        If bStudyNotDistributed = True Then
            Select Case sStudyStatus
            Case eSMStudyStatusNotDist.ssndInPrep 'Is = 1 'Status of in Prep with roles, not distributed
                sStudyStatusIcon = msSTUDY_INPREP_WARNING_ICON
                sToolTip = "Site " & sSiteCode & " is not participating in Study " & sStudyName
            Case eSMStudyStatusNotDist.ssndOpen 'Is = 2 'Status of open, with roles, not distributed
                sStudyStatusIcon = msSTUDY_OPEN_WARNING_ICON
                sToolTip = "Site " & sSiteCode & " is not participating in Study " & sStudyName
            Case eSMStudyStatusNotDist.ssndClosedToFollowup, eSMStudyStatusNotDist.ssndClosedToRecruit _
                        , eSMStudyStatusNotDist.ssndSuspended ' 3 Or 4 or 5, status closed with roles, not distributed
                sStudyStatusIcon = msSTUDY_CLOSED_WARNING_ICON
                sToolTip = "Site " & sSiteCode & " is not participating in Study " & sStudyName
            End Select
        Else
        
            'Status 1 to 5 get from DB, others set by being passed in
            Select Case sStudyStatus
            Case eSMStudyStatusDist.ssdInPrep 'Is = 1 ' study in prep
                sStudyStatusIcon = msSTUDY_INPREP_ICON
                sToolTip = "Study " & sStudyName & " is in preparation"
            Case eSMStudyStatusDist.ssdOpen 'Is = 2 'study open
                sStudyStatusIcon = msSTUDY_OPEN_ICON
                sToolTip = "Study " & sStudyName & " is open"
            Case eSMStudyStatusDist.ssdClosedToRecruit 'Is = 3 'study closed to recruitment
                sStudyStatusIcon = msSTUDY_CLOSED_ICON
                sToolTip = "Study " & sStudyName & " is closed to recruitment"
            Case eSMStudyStatusDist.ssdClosedToFollowup 'Is = 4 'closed to follow up
                sStudyStatusIcon = msSTUDY_CLOSED_ICON
                sToolTip = "Study " & sStudyName & " is closed to follow up"
            Case eSMStudyStatusDist.ssdSuspended 'Is = 5 suspended
                sStudyStatusIcon = msSTUDY_CLOSED_ICON
                sToolTip = "Study " & sStudyName & " is suspended"
            Case eSMStudyStatusDist.ssdDistributedNoRoles 'Is = 6 'distribute with no roles
                sStudyStatusIcon = msSTUDY_DIST_NOROLES_ICON
                sToolTip = sStudyName & " - " & sSiteCode & " combination has no user roles assigned"
            End Select
        End If
        
        If Not DoesNodeExist(sStudyNodeKey) Then
            Set nodx = trvSysAdmin.Nodes.Add(sSiteNodeKey, tvwChild, _
                sStudyNodeKey, sStudyName, sStudyStatusIcon)
            nodx.Tag = msSTUDY_NODE & msSEPARATOR & sToolTip
        End If
    
Exit Sub
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "frmSysAdminTreeView.BuildSiteStudiesNodes"
End Sub

'---------------------------------------------------------------------
Private Sub BuildUserNodes()
'---------------------------------------------------------------------
'REM 07/10/02
'Build the Users nodes
'---------------------------------------------------------------------
Dim nodx As Node
Dim vUsers As Variant
Dim sUsername As String
Dim sUserStatusIcon As String
Dim nUserEnabled As Integer
Dim i As Integer
Dim sToolTip As String
Dim bUserLockedOut As Boolean
Dim nSysAdmin As Integer
Dim lPasswordRetries As Long

    On Error GoTo ErrLabel

    'Create the Users node
    sToolTip = msUSER_NODE_TEXT
    Set nodx = trvSysAdmin.Nodes.Add(, tvwLast, GetUserNodeKey(msUSER_NODE_LABEL), msUSER_NODE_TEXT, msUSERS_ICON)
        nodx.Tag = msUSERS_NODE & msSEPARATOR & sToolTip
        
    'TA 19/1/2006 - for performance password retries is no calculated out of the loop
    lPasswordRetries = PasswordRetries
    
    'get all users in the security database
    vUsers = Users
    If Not IsNull(vUsers) Then
        'add all the individual users to the Users node
        For i = 0 To UBound(vUsers, 2)
        
            nUserEnabled = vUsers(2, i)
            sUsername = vUsers(0, i)
            bUserLockedOut = UserLockout(CLng(vUsers(4, i)), lPasswordRetries)
            nSysAdmin = vUsers(3, i)
            
            'set user status icon
            If (nUserEnabled = 0) And (bUserLockedOut = True) Then
                sUserStatusIcon = msUSER_INACTIVE_ICON
                sToolTip = "User " & sUsername & " has been disabled and locked out"
            ElseIf bUserLockedOut Then
                sUserStatusIcon = msUSER_INACTIVE_ICON
                sToolTip = "User " & sUsername & " has been locked out"
            ElseIf (nUserEnabled = 0) And (nSysAdmin = 0) Then
                sUserStatusIcon = msUSER_INACTIVE_ICON
                sToolTip = "User " & sUsername & " has been disabled"
            ElseIf (nUserEnabled = 0) And (nSysAdmin = 1) Then
                sUserStatusIcon = msUSER_INACTIVE_ICON
                sToolTip = "System Administrator " & sUsername & " has been disabled"
            ElseIf (nUserEnabled = 1) And (nSysAdmin = 0) Then
                sUserStatusIcon = msUSER_ACTIVE_ICON
                sToolTip = "User " & sUsername & " is active"
            ElseIf (nUserEnabled = 1) And (nSysAdmin = 1) Then
                sUserStatusIcon = msUSER_ACTIVE_SYSADMIN_ICON
                sToolTip = "System Administrator " & sUsername & " is active"
            End If
            
            'add each user node
            Set nodx = trvSysAdmin.Nodes.Add(GetUserNodeKey(msUSER_NODE_LABEL), tvwChild, _
                GetUserNodeKey(sUsername), sUsername, sUserStatusIcon)
            nodx.Tag = msUSER_NODE & msSEPARATOR & sToolTip
            
        Next
        
        'expand the node
        nodx.EnsureVisible
    End If
    
Exit Sub
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "frmSysAdminTreeView.BuildUserNodes"
End Sub

'---------------------------------------------------------------------
Private Function UserLockout(lFailedAttempts As Long, lPasswordRetries As Long) As Boolean
'---------------------------------------------------------------------
'REM 287/10/02
'Check if a user is locked out
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsUser As ADODB.Recordset
Dim rsPasswords As ADODB.Recordset

    On Error GoTo ErrHandler
       
    If (lPasswordRetries <> 0) And (lFailedAttempts >= lPasswordRetries) Then
        UserLockout = True
    Else
        UserLockout = False
    End If
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmSysAdminTreeView.UserLockout"
End Function

'---------------------------------------------------------------------
Private Sub BuildRoleNodes()
'---------------------------------------------------------------------
'REM 07/10/02
'Build the Roles nodes
'---------------------------------------------------------------------
Dim nodx As Node
Dim vRoles As Variant
Dim i As Integer
Dim sRoleCode As String
Dim sRoleStatusIcon As String
Dim sToolTip As String


    On Error GoTo ErrLabel

    'Create the Roles node
    sToolTip = msROLES_NODE_TEXT
    Set nodx = trvSysAdmin.Nodes.Add(, tvwLast, GetRoleNodeKey(msROLES_NODE_LABEL), msROLES_NODE_TEXT, msROLES_ICON)
        nodx.Tag = msROLES_NODE & msSEPARATOR & sToolTip
    
    'get all the roles from the security database
    vRoles = Roles
    If Not IsNull(vRoles) Then
        For i = 0 To UBound(vRoles, 2)
            
            sRoleStatusIcon = msROLE_ICON
            
            sRoleCode = vRoles(0, i)
            
            sToolTip = "Role " & sRoleCode
            
            'add each role node
            Set nodx = trvSysAdmin.Nodes.Add(GetRoleNodeKey(msROLES_NODE_LABEL), tvwChild, _
                GetRoleNodeKey(sRoleCode), sRoleCode, sRoleStatusIcon)
            nodx.Tag = msROLE_NODE & msSEPARATOR & sToolTip
            
        Next
        
        'expand the node
        nodx.EnsureVisible
    End If
    
Exit Sub
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "frmSysAdminTreeView.BuildRoleNodes"
End Sub

'---------------------------------------------------------------------
Private Function GetIdFromSelectedItemKey(lGetId As eIdType) As String
'---------------------------------------------------------------------
'REM 03/10/02
'returns the specified part of the key for the selected node
'Key: DatabaseCode|StudyId|StudyName|SiteCode|UserName|RoleCode
'---------------------------------------------------------------------
Dim skey As String
Dim vkey As Variant
Dim lId As Long

    On Error GoTo ErrHandler

    'If there are no items selected in the tree-view then return 0
    If trvSysAdmin.SelectedItem Is Nothing Then
        GetIdFromSelectedItemKey = ""
    Else 'get the id from the key depending on eIdType
        'the key of the node selected
        skey = trvSysAdmin.SelectedItem.Key
        ''read the id's into an array
        vkey = Split(skey, msSEPARATOR)
        'get specific Id
        'if its the study Id then need to check if its = "", if so return "-1"
        If lGetId = StudyId Then
            If vkey(lGetId) = "" Then
                GetIdFromSelectedItemKey = "-1"
            Else
                GetIdFromSelectedItemKey = vkey(lGetId)
            End If
        Else
            GetIdFromSelectedItemKey = vkey(lGetId)
        End If
        
    End If

Exit Function

ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDatalist.GetIdFromSelectedItemKey"
End Function

'---------------------------------------------------------------------
Private Function SelectedNodeToolTip(nodeX As Node) As String
'---------------------------------------------------------------------
'REM 16/10/02
'returns the ToolTip for the selected node
'---------------------------------------------------------------------
Dim vTag As Variant
Dim sNodeTag As String

    sNodeTag = nodeX.Tag
    vTag = Split(sNodeTag, msSEPARATOR)
    
    SelectedNodeToolTip = vTag(1)
    

End Function

'---------------------------------------------------------------------
Private Function SelectedNodeTag(Optional sNTag As String = "") As eSMNodeTag
'---------------------------------------------------------------------
' REM 03/10/02
' Returns the Tag for a selected node in the tree view.
'---------------------------------------------------------------------
Dim sNodeTag As String
Dim vTag As Variant
Dim sTag As String


    On Error GoTo ErrHandler
    
    'if a tag has not been passed in then use th eselected node tag
    If sNTag = "" Then
        sNodeTag = trvSysAdmin.SelectedItem.Tag
    Else
        sNodeTag = sNTag
    End If
    
    'split the tag as it also contains the tool tip
    vTag = Split(sNodeTag, msSEPARATOR)
    'return the first part as this is the tag
    sTag = vTag(0)
    
    Select Case sTag
    Case "U" 'the User node tag
        SelectedNodeTag = eSMNodeTag.UserTag
    Case "Us" ' the Users node tag
        SelectedNodeTag = eSMNodeTag.UsersTag
    Case "D" ' the database tag
        SelectedNodeTag = eSMNodeTag.DatabaseTag
    Case "DD" 'disconnected database tag
        SelectedNodeTag = eSMNodeTag.DisconnectedDB
    Case "UG" 'Database to be upgraded tag
        SelectedNodeTag = eSMNodeTag.Upgrade
    Case "Ds" ' the databases tag
        SelectedNodeTag = eSMNodeTag.DatabasesTag
    Case "St" ' the study tag
        SelectedNodeTag = eSMNodeTag.StudyTag
    Case "Sts" ' the studies tag
        SelectedNodeTag = eSMNodeTag.StudiesTag
    Case "Si" ' the site tage
        SelectedNodeTag = eSMNodeTag.SiteTag
    Case "Sis" ' the sites tag
        SelectedNodeTag = eSMNodeTag.SitesTag
    Case "R" ' the role tag
        SelectedNodeTag = eSMNodeTag.RoleTag
    Case "Rs" ' the roles tag
        SelectedNodeTag = eSMNodeTag.RolesTag
    Case Else
        GoTo ErrHandler
    End Select

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDatalist.SelectedNodeTag"
End Function

'---------------------------------------------------------------------
Private Function GetDatabaseNodeKey(sDatabaseCode As String, Optional sUsername As String = "") As String
'---------------------------------------------------------------------
'REM 03/10/02
'create the database node key
'---------------------------------------------------------------------
    
    If sUsername = "" Then
        GetDatabaseNodeKey = sDatabaseCode & msSEPARATOR & msSEPARATOR & msSEPARATOR & msSEPARATOR & msSEPARATOR
    Else
        GetDatabaseNodeKey = sDatabaseCode & msSEPARATOR & msSEPARATOR & msSEPARATOR & msSEPARATOR & sUsername & msSEPARATOR
    End If
    
End Function

'---------------------------------------------------------------------
Private Function GetStudyNodeKey(sNodeTag As String, sDatabaseCode As String, ByVal lStudyId As Long, sStudyName As String, _
                                Optional sSiteCode As String = "") As String
'---------------------------------------------------------------------
'REM 03/10/02
'create the study node key
'---------------------------------------------------------------------

    If sSiteCode = "" Then
        GetStudyNodeKey = sDatabaseCode & msSEPARATOR & lStudyId & msSEPARATOR & sStudyName & msSEPARATOR & msSEPARATOR & msSEPARATOR & msSEPARATOR & sNodeTag
    Else
        GetStudyNodeKey = sDatabaseCode & msSEPARATOR & lStudyId & msSEPARATOR & sStudyName & msSEPARATOR & sSiteCode & msSEPARATOR & msSEPARATOR & msSEPARATOR & sNodeTag
    End If


End Function

'---------------------------------------------------------------------
Private Function GetSiteNodeKey(sNodeTag As String, sDatabaseCode As String, ByVal sSiteCode As String, Optional lStudyId As Long = -1, Optional sStudyName As String = "") As String
'---------------------------------------------------------------------
'REM 03/10/02
'create the site node key
'---------------------------------------------------------------------

    If lStudyId = -1 Then
        GetSiteNodeKey = sDatabaseCode & msSEPARATOR & msSEPARATOR & msSEPARATOR & sSiteCode & msSEPARATOR & msSEPARATOR & msSEPARATOR & sNodeTag
    Else
        GetSiteNodeKey = sDatabaseCode & msSEPARATOR & lStudyId & msSEPARATOR & sStudyName & msSEPARATOR & sSiteCode & msSEPARATOR & msSEPARATOR & msSEPARATOR & sNodeTag
    End If


End Function

'---------------------------------------------------------------------
Private Function GetUserNodeKey(sUsername As String) As String
'---------------------------------------------------------------------
'REM 03/10/02
'create the user node key
'---------------------------------------------------------------------

    GetUserNodeKey = msSEPARATOR & msSEPARATOR & msSEPARATOR & msSEPARATOR & sUsername & msSEPARATOR


End Function

'---------------------------------------------------------------------
Private Function GetRoleNodeKey(sRoleCode As String) As String
'---------------------------------------------------------------------
'REM 03/10/02
'Creates the role node key
'---------------------------------------------------------------------

    GetRoleNodeKey = msSEPARATOR & msSEPARATOR & msSEPARATOR & msSEPARATOR & msSEPARATOR & sRoleCode


End Function

'---------------------------------------------------------------------
Private Sub trvSysAdmin_Expand(ByVal Node As MSComctlLib.Node)
'---------------------------------------------------------------------
Dim sDatabaseCode As String
    
    'select the node being expanded
    Node.Selected = True
    DoEvents
    If (Node.Image = msDB_CONNECTED_ICON) Or (Node.Image = msDB_LOGGED_IN_ICON) Then
        'get database code from node key
        sDatabaseCode = GetIdFromSelectedItemKey(DatabaseCode)
        'remove dummy nodes before building proper study site nodes
        trvSysAdmin.Nodes.Remove (GetStudyNodeKey(msSTUDIES_NODE, sDatabaseCode, -1, ""))
        trvSysAdmin.Nodes.Remove (GetSiteNodeKey(msSITES_NODE, sDatabaseCode, ""))
        DoEvents
        'build the study site nodes for given database
        Call BuildDatabaseStudiesSites(sDatabaseCode)
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub trvSysAdmin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
'MM 16/10/02
'
'---------------------------------------------------------------------
Dim nodeX As Node

    Set nodeX = trvSysAdmin.HitTest(X, Y)
    If Not nodeX Is Nothing Then
        trvSysAdmin.ToolTipText = SelectedNodeToolTip(nodeX)
    Else
        trvSysAdmin.ToolTipText = ""
    End If
    
    Call RestartSystemIdleTimer

End Sub

'---------------------------------------------------------------------
Private Sub cmdSep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
'MM 16/10/02
'---------------------------------------------------------------------
    
    msglStartX = X

End Sub

'---------------------------------------------------------------------
Private Sub cmdSep_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
'MM 16/10/02
'---------------------------------------------------------------------

Dim sglProposedWidth As Single

    If msglStartX > -1 Then
        sglProposedWidth = cmdSep.Left + X - msglStartX + cmdSep.Width
        If sglProposedWidth < cmdSep.Width Then
            sglProposedWidth = cmdSep.Width
        End If
        RaiseEvent Resize(sglProposedWidth)
        Me.Width = sglProposedWidth
    End If

End Sub

'---------------------------------------------------------------------
Private Sub cmdSep_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    msglStartX = -1

End Sub

'---------------------------------------------------------------------
Private Sub trvSysAdmin_NodeClick(ByVal Node As MSComctlLib.Node)
'---------------------------------------------------------------------
'REM 17/10/02
'Returns the parameters for a selected node
' MLM 20/06/05: bug 2500: tidied this routine and stopped db-specific menu items from being
'   enabled when the db isn't known.
'---------------------------------------------------------------------
'Dim sVersion As String
'Dim sSubVersion As String
Dim enTag As eSMNodeTag
'Dim bAccess As Boolean

    On Error GoTo ErrHandler
    
    enTag = SelectedNodeTag
    RaiseEvent SelectedNode(enTag, GetIdFromSelectedItemKey(DatabaseCode), CLng(GetIdFromSelectedItemKey(StudyId)), GetIdFromSelectedItemKey(StudyName), GetIdFromSelectedItemKey(SiteCode), GetIdFromSelectedItemKey(UserName), GetIdFromSelectedItemKey(RoleCode))
    
    'ASH 24/1/2003 Disable unregisterdatabase option if the selected database
    'node is that of the database currently logged into.
    'NB: had to be done here 'cos menu items are enabled/disabled based on user rights
    'REM 02/04/03 - Added check for GetIdFromSelectedItemKey(DatabaseCode) = "", so if no database selected can't unregister database logged into by a mistake
    If (goUser.Database.DatabaseCode = GetIdFromSelectedItemKey(DatabaseCode)) Or (GetIdFromSelectedItemKey(DatabaseCode) = "") Then
        frmMenu.mnuDunRegisterDatabase.Enabled = False
    End If

    'REM 18/06/03 - Check to see when to enable the upgrade database menu item
    If GetIdFromSelectedItemKey(DatabaseCode) <> "" Then 'check that a database node has been clicked
'        'check database version
'        Call DatabaseVersionSubVersion(GetIdFromSelectedItemKey(DatabaseCode), sVersion, sSubVersion, bAccess)
'        msDatabaseVersion = sVersion & "." & sSubVersion
'
'        If (sVersion = goUser.SecurityDBVersion) And (sSubVersion = goUser.SecurityDBSubVersion) Then 'current version then disable menu
'            frmMenu.mnuUpgradeDatabase.Enabled = False
'        Else 'enable menu for upgrade
'            frmMenu.mnuUpgradeDatabase.Enabled = True
'        End If
        Select Case enTag
            Case eSMNodeTag.DisconnectedDB
                frmMenu.mnuDLockAdministration.Enabled = False
                frmMenu.mnuDTimezone.Enabled = False
                frmMenu.mnuUpgradeDatabase.Enabled = False
            Case eSMNodeTag.Upgrade
                frmMenu.mnuDLockAdministration.Enabled = False
                frmMenu.mnuDTimezone.Enabled = False
                frmMenu.mnuUpgradeDatabase.Enabled = True
            Case eSMNodeTag.DatabaseTag
                frmMenu.mnuDLockAdministration.Enabled = True
                frmMenu.mnuDTimezone.Enabled = True
                frmMenu.mnuUpgradeDatabase.Enabled = False
        End Select
    Else 'database node not clicked
        frmMenu.mnuDLockAdministration.Enabled = False
        frmMenu.mnuDTimezone.Enabled = False
        frmMenu.mnuUpgradeDatabase.Enabled = False
    End If
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDatalist.trvSysAdmin_NodeClick"
End Sub

'---------------------------------------------------------------------
Public Property Get DatabaseVersion() As String
'---------------------------------------------------------------------
'Database version of selected node
'---------------------------------------------------------------------

    DatabaseVersion = msDatabaseVersion

End Property

'---------------------------------------------------------------------
Private Sub trvSysAdmin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
'REM 17/10/02
'Display the right click popup menu on the tree view
'---------------------------------------------------------------------
Dim nodeX As Node
Dim sVersion As String
Dim sSubVersion As String

    On Error GoTo ErrHandler

    'set the node to the one currently under the mouse pointer
    Set nodeX = trvSysAdmin.HitTest(X, Y)
    If Not nodeX Is Nothing Then
        'check for right click
        If Button = vbRightButton Then
            'display the popup menu
            nodeX.Selected = True
            Call trvSysAdmin_NodeClick(nodeX)
            'ASH 9/12/2002 Commented out to fix bug 322 SM
            'RaiseEvent SelectedNode(SelectedNodeTag, GetIdFromSelectedItemKey(DatabaseCode), CLng(GetIdFromSelectedItemKey(StudyId)), GetIdFromSelectedItemKey(StudyName), GetIdFromSelectedItemKey(SiteCode), GetIdFromSelectedItemKey(UserName), GetIdFromSelectedItemKey(RoleCode))
            Call DisplayPopUpMenu(nodeX)
        End If
        
    End If
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDatalist.trvSysAdmin_MouseDown"
End Sub

'---------------------------------------------------------------------
Private Sub DisplayPopUpMenu(nodeX As Node)
'---------------------------------------------------------------------
'REM 17/10/02
'Display the pop up menu with the correct menu items displayed and enabled
'ASH 2/12/2002 Added correct gsFnUnRegisterDatabase to repalce gsFnRegisterDB for
'unregistering of database
'REM 12/03/04 - Added condition compilation arguments so certain menu items are not displayed in the Desktop edition of MACRO
'---------------------------------------------------------------------
Dim sMenuItemSelected As String
Dim oMenuItems As clsMenuItems
Dim Node As Node
Dim sTag As String

    On Error GoTo ErrHandler

    Set oMenuItems = New clsMenuItems

    sTag = nodeX.Tag

    Select Case SelectedNodeTag(sTag)
    Case eSMNodeTag.DatabasesTag
    
    #If DESKTOP <> 1 Then
        Call oMenuItems.Add("CREATEDB", "Create Database", goUser.CheckPermission(gsFnCreateDB))
        Call oMenuItems.Add("REGDB", "Register Database", goUser.CheckPermission(gsFnRegisterDB))
    #End If
    
    Case eSMNodeTag.Upgrade
        Call oMenuItems.Add("UPGRADEDB", "Upgrade Database", True)
        
    Case eSMNodeTag.DatabaseTag
    
    #If DESKTOP <> 1 Then
        Call oMenuItems.Add("SETDBPSWD", "Change Database Password...", goUser.CheckPermission(gsFnChangePassword))
        'ASH 24/1/2003 DO NOT ALLOW UNREGISTERING ON DATABASE CURRENTLY CONNECTED TO /IN USE
        If goUser.Database.DatabaseCode <> GetIdFromSelectedItemKey(DatabaseCode) Then
            Call oMenuItems.Add("UNREGDB", "Unregister Database", goUser.CheckPermission(gsFnUnRegisterDatabase))
        End If
    #End If
    
        Call oMenuItems.Add("LCKADMIN", "Lock Administration", goUser.CheckPermission(gsFnRemoveOwnLocks) Or goUser.CheckPermission(gsFnRemoveAllLocks))
        Call oMenuItems.Add("EDITVIEWROLE", "New User Role", goUser.CheckPermission(gsFnMaintRole))
        Call oMenuItems.Add("LOGDETAILS", "View System Log", goUser.CheckPermission(gsFnViewSystemLog))
        'ASH 8/1/2003 Only show menus if selected database is the same as logged-in database
            If goUser.Database.DatabaseCode = GetIdFromSelectedItemKey(DatabaseCode) Then
                Call oMenuItems.AddSeparator
                Call oMenuItems.Add("SITEADMIN", "Site Administration...", goUser.CheckPermission(gsFnCreateSite))
                Call oMenuItems.Add("STUDYSITEADMIN", "Study Site Administration...", goUser.CheckPermission(gsFnCreateSite))
                Call oMenuItems.Add("LABSITEADMIN", "Laboratory Site Administration...", goUser.CheckPermission(gsFnCreateSite))
                Call oMenuItems.AddSeparator
                Call oMenuItems.Add("EXPSTUDYDEF", "Export Study Definition...", goUser.CheckPermission(gsFnExportStudyDef))
                Call oMenuItems.Add("IMPSTUDYDEF", "Import Study Definition...", goUser.CheckPermission(gsFnImportStudyDef))
                Call oMenuItems.Add("STUDYSTATUS", "Study Status...", goUser.CheckPermission(gsFnChangeTrialStatus))
                Call oMenuItems.Add("GENERATEHTML", "Generate HTML...", True)
                Call oMenuItems.AddSeparator
                Call oMenuItems.Add("EXPSUBDATA", "Export Subjects...", goUser.CheckPermission(gsFnExportPatData))
                Call oMenuItems.Add("IMPSUBDATA", "Import Subjects...", goUser.CheckPermission(gsFnImportPatData))
                Call oMenuItems.AddSeparator
                Call oMenuItems.Add("EXPLAB", "Export Laboratory...", goUser.CheckPermission(gsFnExportLab))
                Call oMenuItems.Add("IMPLAB", "Import Laboratory...", goUser.CheckPermission(gsFnImportLab))
                Call oMenuItems.Add("DISTRIBUTELAB", "Distribute Laboratory...", goUser.CheckPermission(gsFnDistributeLab))

            End If
  
    Case eSMNodeTag.DisconnectedDB
    #If DESKTOP <> 1 Then
        Call oMenuItems.Add("SETDBPSWD", "Change Database Password...", goUser.CheckPermission(gsFnChangePassword))
    #End If
    Case eSMNodeTag.RolesTag
        Call oMenuItems.Add("NEWROLE", "New Role", goUser.CheckPermission(gsFnMaintRole))
        
    Case eSMNodeTag.RoleTag
        Call oMenuItems.Add("EDITVIEWROLE", "View / edit user roles", goUser.CheckPermission(gsFnMaintRole))
        
    Case eSMNodeTag.SitesTag
        Call oMenuItems.Add("EDITVIEWROLE", "View / edit user roles", goUser.CheckPermission(gsFnMaintRole))

    Case eSMNodeTag.SiteTag
        Call oMenuItems.Add("EDITVIEWROLE", "View / edit user roles", goUser.CheckPermission(gsFnMaintRole))

    Case eSMNodeTag.StudiesTag
        Call oMenuItems.Add("EDITVIEWROLE", "View / edit user roles", goUser.CheckPermission(gsFnMaintRole))

    Case eSMNodeTag.StudyTag
        Call oMenuItems.Add("EDITVIEWROLE", "View / edit user roles", goUser.CheckPermission(gsFnMaintRole))
        
    Case eSMNodeTag.UsersTag
        Call oMenuItems.Add("NEWUSER", "New User", goUser.CheckPermission(gsFnCreateNewUser))
        
    Case eSMNodeTag.UserTag
        Call oMenuItems.Add("EDITVIEWROLE", "View / edit user roles", goUser.CheckPermission(gsFnMaintRole))
        
    End Select
    
    'show popup mennu
    sMenuItemSelected = frmMenu.ShowPopUpMenu(oMenuItems)

    Select Case sMenuItemSelected
    Case "CREATEDB" 'create database
    
        Call frmMenu.CreateDatabaseForm
        
    Case "REGDB" 'register a database
        Call frmMenu.RegisterDatabaseForm
        
    Case "UPGRADEDB"
        Call frmMenu.UpgradeDatabase
        
    Case "UNREGDB" 'unregister a database
        Call frmMenu.UnRegisterDatabase(GetIdFromSelectedItemKey(DatabaseCode))
        
    Case "SETDBPSWD" 'set database password
        Call frmMenu.SetDatabasePasswordForm(GetIdFromSelectedItemKey(DatabaseCode))
    
    Case "LCKADMIN" ' lock administration
        Call frmMenu.LockAdministrationForm
    
    Case "LOGDETAILS"
        Call frmMenu.LogDetailsForm(True, GetIdFromSelectedItemKey(DatabaseCode))
    
    Case "EDITVIEWROLE" ' user role
        Call frmMenu.UserRoleForm(False, GetIdFromSelectedItemKey(UserName), GetIdFromSelectedItemKey(DatabaseCode))
        
    Case "NEWROLE" 'new role
        Call frmMenu.RoleManagementForm
    
    Case "NEWUSER" ' new user
        Call frmMenu.NewUserForm(True)
    
    Case "SITEADMIN"
        Call FrmSiteAdmin.Display(GetIdFromSelectedItemKey(DatabaseCode), nodeX)
    
    Case "STUDYSITEADMIN"
        Call frmTrialSiteAdminVersioning.Display(GetIdFromSelectedItemKey(DatabaseCode), DisplayTrialsBySite)
    
    Case "LABSITEADMIN"
        Call frmTrialSiteAdmin.Display(GetIdFromSelectedItemKey(DatabaseCode), DisplaySitesByLab)
    
    Case "EXPSTUDYDEF"
        Call frmExportStudyDefinition.Display
    
    Case "IMPSTUDYDEF"
        Call frmImportStudyDefinition.Display
    
    Case "STUDYSTATUS"
        Call frmTrialStatus.Display(GetIdFromSelectedItemKey(DatabaseCode))
    
    Case "EXPSUBDATA"
        Call frmExportPatientData.Display
    
    Case "IMPSUBDATA"
         Call frmImportPatientData.Display(GetIdFromSelectedItemKey(DatabaseCode))
    
    Case "EXPLAB"
        Call frmExportLab.Display(False, GetIdFromSelectedItemKey(DatabaseCode))
    
    Case "IMPLAB"
        Call frmImportPatientData.Display(GetIdFromSelectedItemKey(DatabaseCode), True)
    
    Case "DISTRIBUTELAB"
        Call frmExportLab.Display(True, GetIdFromSelectedItemKey(DatabaseCode))
    
    Case "GENERATEHTML"
        Call frmGenerateHTML.Display
    End Select
    
    Set oMenuItems = Nothing
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDatalist.DisplayPopUpMenu"
End Sub


