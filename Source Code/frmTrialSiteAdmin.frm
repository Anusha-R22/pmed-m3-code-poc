VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTrialSiteAdmin 
   Caption         =   "Study Site Administration"
   ClientHeight    =   5085
   ClientLeft      =   8310
   ClientTop       =   5385
   ClientWidth     =   7905
   Icon            =   "frmTrialSiteAdmin.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   7905
   Tag             =   "KeepBottomRight"
   Begin VB.CommandButton cmdAll 
      Caption         =   "&Clear All"
      Height          =   375
      Index           =   1
      Left            =   1380
      TabIndex        =   6
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "&Select All"
      Height          =   375
      Index           =   0
      Left            =   60
      TabIndex        =   5
      Top             =   4680
      Width           =   1215
   End
   Begin VB.OptionButton optSitesbyThing 
      Caption         =   "Sites by Laboratory"
      Height          =   255
      Left            =   4740
      TabIndex        =   2
      Top             =   60
      Value           =   -1  'True
      Width           =   2355
   End
   Begin VB.OptionButton optThingsbySite 
      Caption         =   "Laboratories by Site"
      Height          =   255
      Left            =   4740
      TabIndex        =   3
      Top             =   420
      Width           =   2355
   End
   Begin VB.Frame fraCombo 
      Caption         =   "Studies"
      Height          =   735
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   4575
      Begin VB.ComboBox cboSelect 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Tag             =   "KeepBottomRight"
      Top             =   4680
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwList 
      Height          =   3435
      Left            =   180
      TabIndex        =   4
      Tag             =   "resize"
      Top             =   1020
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   6059
      View            =   1
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Site"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "TrialSite"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame fraSites 
      Height          =   3795
      Left            =   60
      TabIndex        =   8
      Top             =   780
      Width           =   7755
   End
   Begin MSComctlLib.ImageList imglistIcons 
      Left            =   5280
      Top             =   4680
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
            Picture         =   "frmTrialSiteAdmin.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrialSiteAdmin.frx":075C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTrialSiteAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       frmTrialSiteAdmin.frm
'   Author:     Joanne Lau, March 1998
'   Purpose:    Maintain cross-reference between trials and sites; and labs and sites
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
'   Revisions:
'   1   Joanne Lau         30/04/98
'   2   Joanne Lau         30/04/98
'   3   Joanne Lau         8/05/98
'   4   Andrew Newbigging  17/07/98
'       Mo Morris           10/12/98    SR 633
'       RefreshList changed. "K|" is added to the key because the site might be numeric
'       and listviews can't have numeric Keys
'       Andrew Newbigging   18/2/99
'       Reference to MSMQ removed
'   5   Paul Norris         21/07/99    Changed form to view sites for a trial and trials for a site
'   6   Paul Norris         21/07/99    Replaced resizing code with resizing object.
'   7   Paul Norris         07/09/99    Upgrade database access code from DAO to ADO
'   8   PN                  15/09/99    Changed call to ADODBConnection() to mconMACRO()
'   PN  21/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   WillC 10/11/99  Added Error handlers
'   Mo 13/12/99     Id's from integer to Long
'   NCJ 14 Dec 99 - Check user's access rights when adding sites/trials
'                   (NB Unresolved problems with extra click events
'                       when showing message boxes)
'                   Added OK button
'   NCJ 21 Dec 99 - Changed "trial" to "study" where it occurs in captions, messages etc.
'   TA 08/05/2000   subclassing removed and resizing now done manually
'   TA 29/09/2000   This form now lets the user link sites to labs
'   TA 06/10/2000:  changed SQL so that lab code not lab id is used and
'                       and allow users to switch view while still in form
'   TA 16/10/2000:  Code added to deal with Site/User administration
'                        and autoselecting the text in the dropdown when loading the form
'   TA 16/10/2000:  Select All / Clear All buttons added
'--------------------------------------------------------------------------------

Option Explicit
Option Compare Binary
Option Base 0

Private Const m_ICON_UNCHECKED = 1
Private Const m_ICON_CHECKED = 2

'min form height and width
Private mlHeight As Long
Private mlWidth As Long


Private meDisplay As eDisplayType
Private mlSelectedTrialId As Long
Private msComboSelectedText As String
'ASH 11/12/2002
Private oDatabase As MACROUserBS30.Database
Private bLoad As Boolean
Private sConnectionString As String
Private sMessage As String
Private mconMACRO As ADODB.Connection
Private msDatabase As String

'---------------------------------------------------------------------
Private Sub cmdAll_Click(Index As Integer)
'---------------------------------------------------------------------
' select/clear all items
'---------------------------------------------------------------------
Dim oListItem As MSComctlLib.ListItem

    HourglassOn
    For Each oListItem In lvwList.ListItems
        If (oListItem.SmallIcon = m_ICON_UNCHECKED And Index = 0) Or (oListItem.SmallIcon = m_ICON_CHECKED And Index = 1) Then
            'if unchecked and index is 0 (Select) OR checked and index = 1 (Clear) then click it
            lvwList_ItemClick oListItem
        End If
    Next

    HourglassOff

End Sub

'---------------------------------------------------------------------
Private Sub cmdOK_Click()
'---------------------------------------------------------------------
' OK button added - NCJ 14/12/99
'---------------------------------------------------------------------

    Unload Me
    
End Sub

'--------------------------------------------------------------------
Public Function Display(ByVal sDatabase As String, eType As eDisplayType, Optional sName As String = "")
'---------------------------------------------------------------------
' display the form according to the eDisplayType
' sName will be selected in combo if passed through
'---------------------------------------------------------------------

Dim columnHdr As ColumnHeader

    On Error GoTo ErrHandler
    
    Me.Icon = frmMenu.Icon
    
    msDatabase = sDatabase
    meDisplay = eType
    
    mlHeight = 4000
    mlWidth = 4000
    FormCentre Me
    
    With lvwList
        'Initialise Image list
        Dim imgX As ListImage
        Call imglistIcons.ListImages.Add(, , LoadResPicture(gsDATA_ITEM_LABEL, vbResIcon))
        Call imglistIcons.ListImages.Add(, , LoadResPicture(gsTICK_LABEL, vbResIcon))
        lvwList.Icons = imglistIcons
        .SmallIcons = imglistIcons
        .Arrange = lvwColumnLeft
    End With
    
    Call DisplayData

    If meDisplay = DisplayTrialsBySite Or meDisplay = DisplayLabsBySite Or meDisplay = DisplayUsersBySite Then
        optThingsbySite.Value = True
    End If
    
    'set the combo if something passed through
    If sName <> "" Then
        cboSelect.Text = sName
    End If
    Me.Show vbModal
    
Exit Function

ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "Display")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
    
End Function

'---------------------------------------------------------------------
Public Sub LoadCombo()
'---------------------------------------------------------------------

'---------------------------------------------------------------------
Dim rsCombo As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler
    
    Set oDatabase = New MACROUserBS30.Database
    bLoad = oDatabase.Load(SecurityADODBConnection, goUser.UserName, msDatabase, "", False, sMessage)
    sConnectionString = oDatabase.ConnectionString
    Set mconMACRO = New ADODB.Connection
    mconMACRO.Open sConnectionString
    mconMACRO.CursorLocation = adUseClient
    ' PN change 5
    ' this requires two loops since the site data does not hold a numeric primary key
    ' the ItemData property of a combo box must numeric
    cboSelect.Clear
    Set rsCombo = New ADODB.Recordset
    
    With rsCombo
        Select Case meDisplay
            Case DisplaySitesByTrial
                sSQL = "SELECT ClinicalTrialName, ClinicalTrialId from ClinicalTrial WHERE ClinicalTrialId > 0"
                .Open sSQL, mconMACRO, adOpenForwardOnly, adLockReadOnly, adCmdText
            Case DisplayTrialsBySite, DisplayLabsBySite, DisplayUsersBySite
                 sSQL = "SELECT Site FROM Site WHERE SiteStatus = 0"
                .Open sSQL, mconMACRO
           Case DisplaySitesByLab
                sSQL = "SELECT LaboratoryCode from Laboratory"
                .Open sSQL, mconMACRO, adOpenForwardOnly, adLockReadOnly, adCmdText
           Case DisplaySitesByUser
                sSQL = "SELECT UserName FROM MACROUser"
                .Open sSQL, SecurityADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
                End Select
                
                Do While Not .EOF
                    cboSelect.AddItem .Fields(0)
                    If .Fields.Count > 1 Then
                        cboSelect.ItemData(cboSelect.NewIndex) = .Fields(1)
                    End If
                    .MoveNext
                Loop
            
            .Close
        End With
    
    Set rsCombo = Nothing
        
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "LoadCombo")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub cboSelect_Click()
'---------------------------------------------------------------------

'---------------------------------------------------------------------

On Error GoTo ErrHandler

    If Not (mlSelectedTrialId = 0 And cboSelect.Text = "") Then
        If Not cboSelect.ListIndex = -1 Then
            mlSelectedTrialId = cboSelect.ItemData(cboSelect.ListIndex)
            msComboSelectedText = cboSelect.Text
            Call RefreshList
        End If
    End If
        
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cboSelect_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
    
End Sub

'---------------------------------------------------------------------
Public Sub LoadListView()
'---------------------------------------------------------------------

'---------------------------------------------------------------------
Dim rsListview As ADODB.Recordset
Dim sSQL As String
Dim sField As String

    On Error GoTo ErrHandler

    lvwList.ListItems.Clear
    lvwList.View = lvwReport
    
    Select Case meDisplay
    Case DisplayTrialsBySite
        sSQL = "SELECT ClinicalTrialName, ClinicalTrialDescription from ClinicalTrial where clinicaltrialid > 0"
    Case DisplaySitesByTrial, DisplaySitesByLab, DisplaySitesByUser
        sSQL = "SELECT Site, SiteDescription FROM Site WHERE SiteStatus = 0"
    Case DisplayLabsBySite
        sSQL = "SELECT LaboratoryCode, LaboratoryDescription FROM Laboratory"
    Case DisplayUsersBySite
         sSQL = "SELECT UserName, UserName FROM MACROUser"
    End Select

    Set rsListview = New ADODB.Recordset
    With rsListview
        If meDisplay = DisplayUsersBySite Then
            .Open sSQL, SecurityADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        Else
            .Open sSQL, mconMACRO, adOpenForwardOnly, adLockReadOnly, adCmdText
        End If
        Do While Not .EOF
            With lvwList.ListItems.Add(, , .Fields(0), , m_ICON_UNCHECKED)
                '.SubItems(1) = rsListview.Fields(1)
            End With
            .MoveNext
        Loop
        .Close
    End With
    
    Set rsListview = Nothing
        
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "LoadListView")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
    
End Sub

'---------------------------------------------------------------------
Public Sub RefreshList()
'---------------------------------------------------------------------

'---------------------------------------------------------------------
Dim sSQL As String
Dim sThisSite
Dim bCurrentSite As Boolean
Dim rsListview As ADODB.Recordset
Dim skey As String
Dim sField As String

    On Error GoTo ErrHandler

    ' lock the window to prevent updates
    Call LockWindow(lvwList)
    
    With lvwList
        .ListItems.Clear
        .Icons = imglistIcons
        .SmallIcons = imglistIcons
        .SortKey = 0 'sort using the listitem object's text property.
        .SortOrder = lvwAscending
        .Sorted = True
    End With
    
    Select Case meDisplay
    Case DisplayTrialsBySite
        sSQL = "SELECT ClinicalTrialName, ClinicalTrialDescription FROM ClinicalTrial WHERE ClinicalTrialId > 0"
    Case DisplaySitesByTrial, DisplaySitesByLab, DisplaySitesByUser
        sSQL = "SELECT Site, SiteDescription FROM Site WHERE SiteStatus = 0"
    Case DisplayLabsBySite
        sSQL = "SELECT LaboratoryCode, LaboratoryDescription FROM Laboratory"
    Case DisplayUsersBySite
         sSQL = "SELECT UserName, UserName FROM MACROUser"
    End Select

    Set rsListview = New ADODB.Recordset
    With rsListview
        If meDisplay = DisplayUsersBySite Then
            .Open sSQL, SecurityADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        Else
            .Open sSQL, mconMACRO, adOpenForwardOnly, adLockReadOnly, adCmdText
        End If
        While Not .EOF
            ' "K|" is added to the key because the site might be numeric
            ' and listviews can't have numeric Keys
            skey = "K|" & .Fields(0)
            sThisSite = .Fields(0)
            
            ' PN change 5
            bCurrentSite = IsMappedSiteOrTrial(sThisSite)
            
            If bCurrentSite Then   'Display a 'check'. Next to the site
                With lvwList.ListItems.Add(, skey, CStr(.Fields(0)), , m_ICON_CHECKED)
                   ' .SubItems(1) = rsListview.Fields(1)
                    'TA 29/09/2000: tag is id
                    .Tag = Format(rsListview.Fields(0))
                End With
                
            Else                        'Display an empty check box.
                With lvwList.ListItems.Add(, skey, CStr(.Fields(0)), , m_ICON_UNCHECKED)
                   ' .SubItems(1) = rsListview.Fields(1)
                    'TA 29/09/2000: tag is id
                    .Tag = Format(rsListview.Fields(0))
                End With
            End If
    
            .MoveNext
        Wend
        
        .Close
    End With
    
    lvwList.Sorted = True
    ' unlock the window for updates
    Call UnlockWindow
    Set rsListview = Nothing


Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "RefreshList")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
    
End Sub

'---------------------------------------------------------------------
Public Function IsMappedSiteOrTrial(ByVal sTrialOrSite As String) As Boolean
'---------------------------------------------------------------------
    'Returns TRUE if Site exists in trial table.
    'Means that this site is associated with the trial.
'---------------------------------------------------------------------
Dim sSQL As String
Dim sField As String
Dim rsMapped As ADODB.Recordset
    
    On Error GoTo ErrHandler
    'sSQL returns a recordset of all sites used for this trial.
    Select Case meDisplay
    Case DisplaySitesByTrial
        sSQL = "SELECT  TrialSite FROM TrialSite, ClinicalTrial " _
                & "WHERE TrialSite.ClinicalTrialId = ClinicalTrial.ClinicalTrialId" _
                & " AND ClinicalTrialName = '" & RTrim(msComboSelectedText) & "'"
    Case DisplayTrialsBySite
        sSQL = "SELECT ClinicalTrial.clinicaltrialname FROM ClinicalTrial, TrialSite" _
                & " WHERE TrialSite.ClinicalTrialId = ClinicalTrial.ClinicalTrialId" _
                & " AND TrialSite.TrialSite = '" & RTrim(msComboSelectedText) & "'"
    Case DisplayLabsBySite
        sSQL = "SELECT LaboratoryCode FROM SiteLaboratory" _
                & " WHERE Site = '" & RTrim(msComboSelectedText) & "'"
    Case DisplaySitesByLab
        sSQL = "SELECT  SiteLaboratory.Site FROM SiteLaboratory" _
                & " WHERE LaboratoryCode = '" & RTrim(msComboSelectedText) & "'"
    Case DisplayUsersBySite
        sSQL = "SELECT UserName FROM SiteUser" _
                & " WHERE Site = '" & RTrim(msComboSelectedText) & "'"
    Case DisplaySitesByUser
        sSQL = "SELECT  SiteUser.Site FROM SiteUser" _
                & " WHERE UserName = '" & RTrim(msComboSelectedText) & "'"
    End Select
    
    Set rsMapped = New ADODB.Recordset
    With rsMapped

       .Open sSQL, mconMACRO, adOpenForwardOnly, adLockReadOnly, adCmdText

        'Compare each(active) site (ThisSite - in the listview) with the set of sites used for
        'this trial.If a match is found ,return true. The site is 'checked' in the listview.
        IsMappedSiteOrTrial = False
        Do While Not .EOF
            If sTrialOrSite = .Fields(0) Then
                IsMappedSiteOrTrial = True
                Exit Do
            End If
            .MoveNext
        Loop
    
        .Close
    End With
    
    Set rsMapped = Nothing
        
Exit Function
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "IsMappedSiteOrTrial")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
    
End Function

Private Sub Form_Resize()

On Error GoTo ErrHandler

    If Me.WindowState <> vbMinimized Then
        If Me.Height >= mlHeight Then
            fraSites.Height = Me.ScaleHeight - cmdOK.Height - fraCombo.Height - 240
            lvwList.Height = fraSites.Height - 360
            cmdOK.Top = fraSites.Top + fraSites.Height + 120
            '   Bug Fix Macro 2.2 008 RJCW 15/10/2001
            cmdAll(0).Top = cmdOK.Top
            cmdAll(1).Top = cmdOK.Top
        End If
        If Me.Width >= mlWidth Then
            fraSites.Width = Me.ScaleWidth - 120
            lvwList.Width = fraSites.Width - 240
            cmdOK.Left = fraSites.Left + fraSites.Width - cmdOK.Width
        End If
    End If
    
ErrHandler:

End Sub

'---------------------------------------------------------------------
Private Sub UpdateSites(ByVal Item As MSComctlLib.ListItem)
'---------------------------------------------------------------------
' Check the current user is allowed to do this... NCJ 14/12/99
' (We check here in case they have permission
' to EITHER add OR remove trial/sites)
'---------------------------------------------------------------------
Dim sSQL As String
Dim oSiteStatus As ADODB.Recordset
Dim bNewTrialSite As Boolean
Dim sSelectedTrialName As String
Dim oTrialID As ADODB.Recordset
    
    On Error GoTo ErrHandler

    If cboSelect.Text > "" Then
        sSelectedTrialName = Item
        
        'Get Sites for this trial
        sSQL = "SELECT TrialSite FROM TrialSite, ClinicalTrial "
        sSQL = sSQL & "WHERE ClinicalTrial.ClinicalTrialName = '" & sSelectedTrialName & "' "
        sSQL = sSQL & "AND TrialSite.ClinicalTrialId = ClinicalTrial.ClinicalTrialId "
        sSQL = sSQL & "AND TrialSite.TrialSite = '" & msComboSelectedText & "'"
    
        ' PN 07/09/99
        ' upgrade to ado from dao
        Set oSiteStatus = New ADODB.Recordset
        With oSiteStatus
            .Open sSQL, mconMACRO, adOpenForwardOnly, adLockReadOnly, adCmdText
            bNewTrialSite = .EOF
            .Close
        End With
        Set oSiteStatus = Nothing

        ' get the trial ID
        sSQL = "SELECT ClinicalTrialID FROM ClinicalTrial where ClinicalTrialName ='" & sSelectedTrialName & "'"
        Set oTrialID = New ADODB.Recordset
        With oTrialID
            .Open sSQL, mconMACRO, adOpenForwardOnly, adLockReadOnly, adCmdText
            
            'Update the database
            If bNewTrialSite Then
            
                If goUser.CheckPermission(gsFnAddSiteToTrialOrTrialToSite) Then
                    'TA 16/04/2002: Deleted Actual Recruitment field
                    sSQL = "INSERT into TrialSite (ClinicalTrialId, TrialSite) Values (" _
                         & oTrialID.Fields("ClinicalTrialID") & ",'" & msComboSelectedText & "')"
                    mconMACRO.Execute sSQL
                    
                    '!   ATN 18/2/99
                    '   Removed reference to MSMQ
                    CreateStatusMessage ExchangeMessageType.NewTrial, _
                                        .Fields("ClinicalTrialID"), _
                                        sSelectedTrialName, _
                                         msComboSelectedText
                    Item.SmallIcon = m_ICON_CHECKED
                Else
                ' Showing the message box here causes an extra lvwClick event to occur
                ' and the message box is shown twice
'                    MsgBox "You do not have permission to add trials to sites", vbOKOnly, "MACRO Exchange"
'                    Debug.Print "Done message box"
                End If
                
            Else 'It is Active, change it to InActive by deleting the record.
            ' NCJ 14/12/99 Check their access rights
                If goUser.CheckPermission(gsFnRemoveSite) Then
                    'Changed by Mo Morris 18/1/00. '*' removed from Delete SQL statement
                    sSQL = "DELETE FROM TrialSite WHERE ClinicalTrialId = " & .Fields("ClinicalTrialID") & _
                           " AND TrialSite.TrialSite = '" & msComboSelectedText & "'"
                    mconMACRO.Execute sSQL
                    Item.SmallIcon = m_ICON_UNCHECKED
                Else
                ' Showing the message box here causes an extra lvwClick event to occur
                ' and the message box is shown twice
'                    MsgBox "You do not have permission to remove trials from sites", vbOKOnly, "MACRO Exchange"
                End If
            End If
            
            'Call RefreshList
            .Close
            
        End With
        Set oTrialID = Nothing
        
    End If
        
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "UpdateSites")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
    
End Sub
'---------------------------------------------------------------------
Private Sub UpdateTrials(ByVal Item As MSComctlLib.ListItem)
'---------------------------------------------------------------------
' Check the current user is allowed to do this... NCJ 14/12/99
' (We check here in case they have permission
' to EITHER add OR remove trial/sites)
'---------------------------------------------------------------------

Dim sSQL As String
Dim oSiteStatus As ADODB.Recordset
Dim bNewTrialSite As Boolean
Dim sSelectedSiteName As String
    
    On Error GoTo ErrHandler
    
    If cboSelect.Text > "" Then
    
        sSelectedSiteName = Item
        
        'Get Sites for this trial
        sSQL = "SELECT TrialSite FROM TrialSite, ClinicalTrial " _
                & "WHERE TrialSite.TrialSite = '" & sSelectedSiteName & "'" _
                & " AND TrialSite.ClinicalTrialId = ClinicalTrial.ClinicalTrialId" _
                & " AND ClinicalTrial.ClinicalTrialName = '" & msComboSelectedText & "'"
    
        ' PN 07/09/99
        ' upgrade to ado from dao
        Set oSiteStatus = New ADODB.Recordset
        With oSiteStatus
            .Open sSQL, mconMACRO, adOpenForwardOnly, adLockReadOnly, adCmdText
            bNewTrialSite = .EOF
            .Close
        End With

        'Update the database
        If bNewTrialSite Then
            ' NCJ 14/12/99 Check their access rights
            If goUser.CheckPermission(gsFnAddSiteToTrialOrTrialToSite) Then
                'TA 16/04/2002: remove reference to acutal recruitment
                sSQL = "INSERT into TrialSite (ClinicalTrialId, TrialSite) Values (" _
                     & mlSelectedTrialId & ",'" & sSelectedSiteName & "')"
                mconMACRO.Execute sSQL
                
                '   ATN 18/2/99
                '   Removed reference to MSMQ
                CreateStatusMessage ExchangeMessageType.NewTrial, _
                                    mlSelectedTrialId, _
                                    msComboSelectedText, _
                                     sSelectedSiteName
            Item.SmallIcon = m_ICON_CHECKED
            Else
                ' Showing the message box here causes an extra lvwClick event to occur
                ' and the message box is shown twice
'                MsgBox "You do not have permission to add studies to sites", vbOKOnly, "MACRO Exchange"
'                    Debug.Print "Done message box"
            End If
    
        Else 'It is Active, change it to InActive by deleting the record.
            ' NCJ 14/12/99 Check their access rights
            If goUser.CheckPermission(gsFnRemoveSite) Then
                'Changed by Mo Morris 29/11/99. '*' removed from Delete SQL statement
                sSQL = "DELETE FROM TrialSite WHERE ClinicalTrialId = " & mlSelectedTrialId & _
                                             " AND TrialSite.TrialSite = '" & sSelectedSiteName & "'"
                mconMACRO.Execute sSQL
                Item.SmallIcon = m_ICON_UNCHECKED
            Else
                ' Showing the message box here causes an extra lvwClick event to occur
                ' and the message box is shown twice
'                MsgBox "You do not have permission to remove studies from sites", vbOKOnly, "MACRO Exchange"
            End If
            
        End If
        
    End If
        
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "UpdateTrials")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub UpdateLabSite(ByVal Item As MSComctlLib.ListItem)
'---------------------------------------------------------------------
' update SiteLaboratory table according to user selection
'---------------------------------------------------------------------
Dim sSQL As String
Dim sSite As String
Dim sLabCode As String
    
    On Error GoTo ErrHandler
    
    If cboSelect.Text > "" Then
    
        If meDisplay = eDisplayType.DisplayLabsBySite Then
            sSite = cboSelect.Text
            sLabCode = Item.Tag
        Else
            sSite = Item.Tag
            sLabCode = cboSelect.Text
        End If
        
        Select Case Item.SmallIcon
        Case m_ICON_UNCHECKED
            'unchecked to checked - add
            sSQL = "INSERT INTO SiteLaboratory (Site, LaboratoryCode) VALUES ('" & sSite & "','" & sLabCode & "')"
            Item.SmallIcon = m_ICON_CHECKED
        Case m_ICON_CHECKED
            'checked to unchecked - remove
            sSQL = "DELETE FROM SiteLaboratory WHERE Site = '" & sSite & "' AND LaboratoryCode = '" & sLabCode & "'"
            Item.SmallIcon = m_ICON_UNCHECKED
        End Select
        
        mconMACRO.Execute sSQL
    End If
    

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "UpdateLabSite")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
    
    
End Sub

'---------------------------------------------------------------------
Private Sub UpdateUserSite(ByVal Item As MSComctlLib.ListItem)
'---------------------------------------------------------------------
' update SiteLaboratory table according to user selection
'---------------------------------------------------------------------
Dim sSQL As String
Dim sSite As String
Dim sUser As String
    
    On Error GoTo ErrHandler
    
    If cboSelect.Text > "" Then
    
        If meDisplay = eDisplayType.DisplayUsersBySite Then
            sSite = cboSelect.Text
            sUser = Item.Tag
        Else
            sSite = Item.Tag
            sUser = cboSelect.Text
        End If
        
        Select Case Item.SmallIcon
        Case m_ICON_UNCHECKED
            'unchecked to checked - add
            sSQL = "INSERT INTO SiteUser (Site, UserName) VALUES ('" & sSite & "','" & sUser & "')"
            Item.SmallIcon = m_ICON_CHECKED
        Case m_ICON_CHECKED
            'checked to unchecked - remove
            sSQL = "DELETE FROM SiteUser WHERE Site = '" & sSite & "' AND UserName = '" & sUser & "'"
            Item.SmallIcon = m_ICON_UNCHECKED
        End Select
        
        mconMACRO.Execute sSQL
    End If
    

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "UpdateUserSite")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
    
    
End Sub

'---------------------------------------------------------------------
Private Sub lvwList_ItemClick(ByVal Item As MSComctlLib.ListItem)
'---------------------------------------------------------------------
'When an item on the list view is clicked, causing the item to be checked,
'a record is created in the TrialSite table. If the click causes it to be unchecked,
'the corresponding record is deleted from the table.
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    HourglassOn

    Select Case meDisplay
    Case DisplaySitesByTrial: Call UpdateTrials(Item)
    Case DisplayTrialsBySite: Call UpdateSites(Item)
    Case DisplaySitesByLab: Call UpdateLabSite(Item)
    Case DisplayLabsBySite: Call UpdateLabSite(Item)
    Case DisplaySitesByUser: Call UpdateUserSite(Item)
    Case DisplayUsersBySite: Call UpdateUserSite(Item)
    End Select
        
    HourglassOff
        
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "lvwList_ItemClick")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub optThingsbySite_Click()
'---------------------------------------------------------------------
    Call ChangeRound

End Sub

'---------------------------------------------------------------------
Private Sub optSitesbyThing_Click()
'---------------------------------------------------------------------
    Call ChangeRound
    
End Sub

'---------------------------------------------------------------------
Private Sub ChangeRound()
'---------------------------------------------------------------------
' switch between things by Site and Sites by things
'---------------------------------------------------------------------

    

    If optThingsbySite.Value Then
        Select Case meDisplay
        Case DisplaySitesByLab: meDisplay = DisplayLabsBySite
        Case DisplaySitesByTrial: meDisplay = DisplayTrialsBySite
        Case DisplaySitesByUser: meDisplay = DisplayUsersBySite
        End Select
    Else
        Select Case meDisplay
        Case DisplayLabsBySite: meDisplay = DisplaySitesByLab
        Case DisplayTrialsBySite: meDisplay = DisplaySitesByTrial
        Case DisplayUsersBySite: meDisplay = DisplaySitesByUser
        End Select
    End If
    
    Call DisplayData
    
End Sub

'---------------------------------------------------------------------
Private Sub DisplayData()
'---------------------------------------------------------------------
'redisplay all data
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    Call LoadCombo
    
    
    Select Case meDisplay
    Case DisplaySitesByTrial, DisplayTrialsBySite
        Me.Caption = "Study Site Administration " & "[" & goUser.DatabaseCode & "]"
        optThingsbySite.Caption = "Studies by Site"
        optSitesbyThing.Caption = "Sites by Study"
        
    Case DisplaySitesByLab, DisplayLabsBySite
        Me.Caption = "Laboratory Site Administration " & "[" & goUser.DatabaseCode & "]"
        optThingsbySite.Caption = "Laboratories by Site"
        optSitesbyThing.Caption = "Sites by Laboratory"
        
    Case DisplaySitesByUser, DisplayUsersBySite
        Me.Caption = "User Site Administration " & "[" & goUser.DatabaseCode & "]"
        optThingsbySite.Caption = "Users by Site"
        optSitesbyThing.Caption = "Sites by User"
        
    End Select
    

    
    With lvwList
        .ColumnHeaders.Clear
    
        Select Case meDisplay
        Case DisplaySitesByTrial
            fraCombo.Caption = "Study"
            Call .ColumnHeaders.Add(, , "Active sites", 2000, lvwColumnLeft)
        Case DisplayTrialsBySite
            fraCombo.Caption = "Active Site"
            Call .ColumnHeaders.Add(, , "Studies", 2000, lvwColumnLeft)
        Case DisplaySitesByLab
            fraCombo.Caption = "Laboratory"
            Call .ColumnHeaders.Add(, , "Active sites", 2000, lvwColumnLeft)
        Case DisplayLabsBySite
            fraCombo.Caption = "Active Site"
            Call .ColumnHeaders.Add(, , "Laboratories", 2000, lvwColumnLeft)
        Case DisplaySitesByUser
            fraCombo.Caption = "User"
            Call .ColumnHeaders.Add(, , "Active sites", 2000, lvwColumnLeft)
        Case DisplayUsersBySite
            fraCombo.Caption = "Active Site"
            Call .ColumnHeaders.Add(, , "Users", 2000, lvwColumnLeft)
        End Select
        
        '.ColumnHeaders.Add , , "Description", 2000, lvwColumnLeft
        
        .MultiSelect = True
    
        'Sort on first column
        .SortKey = 0 'sort using the listitem object's text property.
        .SortOrder = lvwAscending
        .Sorted = True
        

    End With
    
    Call LoadListView
    
    
   ' NCJ 14/12/99 & TA 17/10/2000
    ' If they are in trial/site admin and they can't add or remove trials/sites then disable the listview
    If meDisplay = eDisplayType.DisplayTrialsBySite Or meDisplay = eDisplayType.DisplaySitesByTrial Then
        lvwList.Enabled = goUser.CheckPermission(gsFnAddSiteToTrialOrTrialToSite) Or goUser.CheckPermission(gsFnRemoveSite)
    End If
    
    If cboSelect.ListCount > 0 Then
        cboSelect.ListIndex = 0
    End If
    
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "DisplayData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
End Sub
