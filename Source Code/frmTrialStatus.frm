VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTrialStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of Studies"
   ClientHeight    =   6825
   ClientLeft      =   570
   ClientTop       =   1395
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6825
   ScaleWidth      =   12090
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraOptions 
      Height          =   6075
      Left            =   9480
      TabIndex        =   6
      Top             =   660
      Width           =   2535
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   5470
         Width           =   2055
      End
      Begin VB.CommandButton cmdNewVersion 
         Caption         =   "Create new version"
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   2055
      End
      Begin VB.Frame fraHistory 
         Caption         =   "View"
         Height          =   1515
         Left            =   120
         TabIndex        =   12
         Top             =   3720
         Width           =   2295
         Begin VB.CommandButton CmdTrialSites 
            Caption         =   "Study Sites"
            Height          =   495
            Left            =   120
            TabIndex        =   14
            Top             =   900
            Width           =   2055
         End
         Begin VB.CommandButton CmdHistory 
            Caption         =   "History"
            Height          =   495
            Left            =   120
            TabIndex        =   13
            Top             =   300
            Width           =   2055
         End
      End
      Begin VB.Frame fraStatus 
         Caption         =   "Status"
         Height          =   2655
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   2295
         Begin VB.CommandButton cmdInPreparation 
            Caption         =   "In preparation"
            Height          =   395
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton CmdCloseFollowUp 
            Caption         =   "Closed to follow up"
            Height          =   395
            Left            =   120
            TabIndex        =   11
            Top             =   2160
            Width           =   2055
         End
         Begin VB.CommandButton CmdCloseRecruitment 
            Caption         =   "Closed to recruitment"
            Height          =   395
            Left            =   120
            TabIndex        =   10
            Top             =   1680
            Width           =   2055
         End
         Begin VB.CommandButton CmdSuspend 
            Caption         =   "Suspended"
            Height          =   395
            Left            =   120
            TabIndex        =   9
            Top             =   1200
            Width           =   2055
         End
         Begin VB.CommandButton CmdOpenTrial 
            Caption         =   "Open"
            Height          =   395
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   2055
         End
      End
   End
   Begin VB.PictureBox picSearch 
      Height          =   540
      Left            =   120
      ScaleHeight     =   480
      ScaleWidth      =   11835
      TabIndex        =   1
      Top             =   108
      Width           =   11895
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   330
         Left            =   6000
         TabIndex        =   5
         Top             =   108
         Width           =   1065
      End
      Begin VB.CommandButton cmdKeywordSearch 
         Caption         =   "Search"
         Height          =   330
         Left            =   4680
         TabIndex        =   4
         Top             =   105
         Width           =   1065
      End
      Begin VB.TextBox txtKeywordSearch 
         Height          =   330
         Left            =   1470
         TabIndex        =   3
         Top             =   105
         Width           =   2955
      End
      Begin VB.Label lblKeywordSearch 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Keyword:"
         Height          =   195
         Left            =   615
         TabIndex        =   2
         Top             =   105
         Width           =   660
      End
   End
   Begin MSComctlLib.ListView lvwTrials 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   10610
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imglistSmallIcons 
      Left            =   2760
      Top             =   948
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imglistLargeIcons 
      Left            =   1992
      Top             =   840
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
End
Attribute VB_Name = "frmTrialStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       frmTrialStatus.frm
'   Author:     Joanne Lau, April 1998
'   Purpose:    Maintain status of trials and distribute new versions.  Creates messages
'   in MSMQ queues for sites associated with changed trial
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
 '  1   Joanne Lau          30/04/98
'   2   Joanne Lau          30/04/98
'   3   Joanne Lau          6/05/98
'   4   Joanne Lau          8/05/98
'   5   Andrew Newbigging   17/07/98
'   6   Andrew Newbigging   22/10/98
'   7   Andrew Newbigging   18/2/99
'       Modified cmdNewVersion so that study definition is NOT stored in MSMQ.
'       Note: MSMQ no longer used anyway.
'       Now uses TrialStatus enumeration.
'   8  PN  10/09/99     Upgrade from DAO to ADO and updated code to conform
'                       to VB standards doc version 1.0
'   9  PN  15/09/99     Changed call to ADODBConnection() to MacroADODBConnection()
'   PN  21/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   PN  29/09/99    Conversion of cmdHistory_Click() to ADO form DAO
'   WillC   10/11/99    Added the Errhandlers
'   Mo 13/12/99     Id's from integer to Long
'   NCJ 14/12/99    Check user's access rights
'                   Added OK button
'   NCJ 21/12/99    Changed "trial" to "study" in messages
'                   Show timestamp as date instead of double
'                   Change TrialStatus to eTrialStatus
'   Mo Morris   13/1/00 SR2662
'               changes in Form_Load and lvwTrials_ItemClick.
'               Enabling/disabling of command buttons within frame fraOptions changed from
'               simply enabling/disabling the frame to actually enabling/disabling the
'               specific controls.
'   Mo Morris   6/3/00  SR2764
'               RefreshTrialList and cmdKeywordSearch_Click now call new private sub DisableCommands
'   WillC 24/5/00 SR3492 Warn the user if they set a staus back to In Prep form Open
'   WillC 24/5/00 Changed
'   TA 16/10/2000: Select Study selected in listview when showing TrialSite admin
'   DPH 17/10/2001 Added calls to FolderExistence routine to create missing folders
'   ATO 01/02/2002 Modified RefreshTrialList routine by adding new sql to display correct
'                  ActualTrialSubjects
'   DPH 03/05/2002 Check added in distribute new version before setting message for sites
'   DPH 29/05/2002 - Changed SQL to disallow inactive sites from distribution in CreateMessageForAllTrialSites
'   DPH 29/09/2006 - Bug 2757 - corrected Keywords appearing under wrong column header for keyword search
'----------------------------------------------------------------------------------------'

Option Base 0
Option Explicit
Option Compare Binary

Private Const msKEY_SEPARATOR = "_"

Private msSelectedTriaStatus As String 'Selected trial status SR3492
Private mnSelectedClinicalTrialId As Long 'SD selected trial in list view.
Private msSelectedClinicalTrialName As String
Private mnClinicalTrialId As Long 'used to store form property-frmSD
'ASH 13/12/2002
Private oDatabase As MACROUserBS30.Database
Private bLoad As Boolean
Private sConnectionString As String
Private sMessage As String
Private mconMACRO As ADODB.Connection
Private msDatabase As String


'---------------------------------------------------------------------
Private Sub cmdInPreparation_Click()
'---------------------------------------------------------------------
' Change status of study to "In Preparation"
'---------------------------------------------------------------------
Dim sMsg As String
    On Error GoTo ErrHandler
    
    
    sMsg = "If you change this study to In Preparation you will be able to make changes to the" & vbCrLf _
            & "Study Definition which could invalidate existing patient data." & vbCrLf _
            & "It is not advisable to make this change." & vbCrLf _
            & "Are you sure you wish to continue?"

'    'WillC 24/5/00 SR3492 Warn the user if they set a staus back to In Prep form Open
    If msSelectedTriaStatus = "Open" Then
        Select Case MsgBox(sMsg, vbCritical + vbOKCancel, "MACRO")
                Case vbOK
                        Call UpdateTrialStatus(ClinicalTrialId, _
                            msSelectedClinicalTrialName, _
                            eTrialStatus.InPreparation)
                        Call ChangeStatus(eTrialStatus.ClosedToFollowUp)
                        Call RefreshTrialList
                Case vbCancel
                    Exit Sub
        End Select
    End If
    
'    Call UpdateTrialStatus(ClinicalTrialId, _
'                            msSelectedClinicalTrialName, _
'                            eTrialStatus.inpreparation)
'    Call ChangeStatus(eTrialStatus.ClosedToFollowUp)
'    Call RefreshTrialList
        
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "cmdInPreparation_Click")
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
Private Sub cmdOK_Click()
'---------------------------------------------------------------------
' OK button added - NCJ 14 Dec 99
'---------------------------------------------------------------------

    Unload Me

End Sub

'----------------------------------------------------------------------------------------'
Public Property Get SelectedClinicalTrialId() As Integer
'----------------------------------------------------------------------------------------'

    SelectedClinicalTrialId = mnSelectedClinicalTrialId

End Property

'----------------------------------------------------------------------------------------'
Public Property Get SelectedClinicalTrialName() As String
'----------------------------------------------------------------------------------------'

    SelectedClinicalTrialName = msSelectedClinicalTrialName

End Property

'****
'----------------------------------------------------------------------------------------'
Public Property Get ClinicalTrialId() As Long '****From frmStudyDefinition****
'----------------------------------------------------------------------------------------'

    ClinicalTrialId = mnSelectedClinicalTrialId ' return value '
    
End Property

'***from frmSD ** replaced arg tmpStr
'----------------------------------------------------------------------------------------'
Public Property Let ClinicalTrialId(mnSelectedClinicalTrialId As Long)
'----------------------------------------------------------------------------------------'

    mnClinicalTrialId = mnSelectedClinicalTrialId ' store value

End Property

'-----------------------------------------------------------
Private Sub RefreshTrialList()
'-----------------------------------------------------------
' SUB: RefreshTrialList
'
' Reads the list of trials from the database and displays
' it in the form.
'
'-----------------------------------------------------------
' REVISIONS
' DPH 03/09/2002 - Add Latest Study Version
'-----------------------------------------------------------
Dim oTrialItem As ListItem
Dim rsTrialList As ADODB.Recordset
Dim rsTotal As ADODB.Recordset
Dim skey As String
Dim sSQL As String
Dim sRecruitmentSQL As String

    On Error GoTo ErrHandler
    
    Me.MousePointer = vbHourglass
    
    'Disable command buttons
    DisableCommands
    
    'Remove existing items
    lvwTrials.ListItems.Clear
    
    ' Get the list of trials, excluding the library
    sSQL = "SELECT * FROM ClinicalTrial WHERE ClinicalTrialId > 0"
    Set rsTrialList = New ADODB.Recordset
    rsTrialList.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Set rsTrialList = New ADODB.Recordset
    rsTrialList.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    ' While the record is not the last record, add a ListItem object.
    While Not rsTrialList.EOF
    
        ' Store the sponsor and trial id in the listitem key
        skey = msKEY_SEPARATOR & rsTrialList![ClinicalTrialId]
        
        sRecruitmentSQL = "SELECT COUNT(*)"
        sRecruitmentSQL = sRecruitmentSQL & " FROM TrialSubject"
        sRecruitmentSQL = sRecruitmentSQL & " WHERE TrialSubject.ClinicalTrialID=" & rsTrialList![ClinicalTrialId]
        
        Set rsTotal = New ADODB.Recordset
        rsTotal.Open sRecruitmentSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
                
        ' add the listitem
        Set oTrialItem = lvwTrials.ListItems.Add(, skey, _
            rsTrialList![ClinicalTrialName], gsTRIAL_LABEL, gsTRIAL_LABEL)
        
        With oTrialItem
            ' look up the description of the trial status
            .SubItems(1) = gsTrialStatus(rsTrialList![statusId])
            ' look up the description of the trial phase
            .SubItems(2) = gsTrialPhase(rsTrialList![PhaseId])
            .SubItems(3) = rsTrialList![ExpectedRecruitment]
            .SubItems(4) = rsTotal.Fields(0).Value
            ' DPH 03/09/2002 - Get Latest Version Number
            .SubItems(5) = GetLatestVersionOfTrialAvailable(rsTrialList![ClinicalTrialId])
            If Not IsNull(rsTrialList![keywords]) Then
                .SubItems(6) = rsTrialList![keywords]
            End If
            
        End With
        
        rsTrialList.MoveNext   ' Move to next record.
    Wend
    
    ' Close the recordset
    rsTrialList.Close
    
    Me.MousePointer = vbDefault
    
    'make sure that no trial is selected initially
    Set lvwTrials.SelectedItem = Nothing
    Set rsTrialList = Nothing
    
    'Disable the Search command and clear down the search text
    cmdKeywordSearch.Enabled = False
    txtKeywordSearch.Text = vbNullString
        
    If lvwTrials.ListItems.Count > 0 Then
        lvwTrials.ListItems(1).Selected = True
        lvwTrials_ItemClick lvwTrials.ListItems(1)
    End If
    
    'REM 21/05/03 - Update User Study/Site permissions
    goUser.ReloadStudySitePermissions
    
    EnableUsersButtons
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "RefreshTrialList")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
        
    
End Sub

'-----------------------------------------------------------
Private Sub CmdCloseFollowUp_Click()
'-----------------------------------------------------------
    
    On Error GoTo ErrHandler

    Call UpdateTrialStatus(ClinicalTrialId, _
                            msSelectedClinicalTrialName, _
                            eTrialStatus.ClosedToFollowUp)
    Call ChangeStatus(eTrialStatus.ClosedToFollowUp)
    Call RefreshTrialList
        
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "CmdCloseFollowUp_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
        
    
End Sub

'-----------------------------------------------------------
Private Sub CmdCloseRecruitment_Click()
'-----------------------------------------------------------
'-----------------------------------------------------------

    On Error GoTo ErrHandler

    Call UpdateTrialStatus(ClinicalTrialId, msSelectedClinicalTrialName, _
                            eTrialStatus.ClosedToRecruitment)
    Call ChangeStatus(eTrialStatus.ClosedToRecruitment)
    Call RefreshTrialList
        
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "CmdCloseRecruitment_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub

'-----------------------------------------------------------
Private Sub cmdHistory_Click()
'-----------------------------------------------------------
' PN 29/09/99   convert routine to return an ADO recordset
'-----------------------------------------------------------
Dim rsTrialHistory As ADODB.Recordset 'mnVersionId
Dim sMsg As String
Dim sVersion As String
Dim sStatus As String
Dim sUser As String
Dim sTime As String
    
   On Error GoTo ErrHandler
   
    
   ' WillC 29/2/00 SR2666 Took out the message box and added a new form.
   Call frmStudyStatusHistory.RefreshMe(mnSelectedClinicalTrialId, msSelectedClinicalTrialName)
    
'    Set rsTrialHistory = gdsTrialHistory(mnSelectedClinicalTrialId)
'    'Mo 13/1/00, spaces changed to tabs for the purpose of lining up the heading line with the data line
'    sMsg = "Version Number" & vbTab & "Study Status" & vbTab & vbTab & "User" & vbTab & "Date/time" & vbCr
'    With rsTrialHistory
'        While Not .EOF
'            sVersion = rsTrialHistory!VersionId
'            sStatus = StatusDescription(rsTrialHistory!statusId)
'            sUser = rsTrialHistory!StudyDefinitionUserId
'            ' NCJ 21/12/99 - Format timestamp from double to date
'            'mo 13/1/00 SR2667, format string changed from dd/mm/yyyy to yyyy/mm/dd hh:mm:ss
'            sTime = Format(CDate(rsTrialHistory!StatusChangedTimestamp), "yyyy/mm/dd hh:mm:ss")
'
'            sMsg = sMsg & sVersion & vbTab & vbTab & sStatus & vbTab
'            sMsg = sMsg & sUser & vbTab & sTime & vbTab & vbNewLine
'            .MoveNext
'        Wend
'        .Close
'
'    End With
'
'    Set rsTrialHistory = Nothing
'
'    sMsg = sMsg & vbNewLine
'    MsgBox sMsg, , "Study Status History"
        
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "CmdHistory_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub

'-----------------------------------------------------------
Public Function StatusDescription(nStatusId As Integer) As String
'-----------------------------------------------------------
On Error GoTo ErrHandler

    Select Case nStatusId
    Case eTrialStatus.InPreparation
        StatusDescription = "In preparation" & vbTab
    Case eTrialStatus.TrialOpen
        StatusDescription = "Open" & vbTab & vbTab
    Case eTrialStatus.ClosedToRecruitment
        StatusDescription = "Closed to recruitment"
    Case eTrialStatus.ClosedToFollowUp
        StatusDescription = "Closed to follow Up" & vbTab
    Case eTrialStatus.Suspended
        StatusDescription = "Suspended" & vbTab
    End Select
        
Exit Function
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "StatusDescription")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Function

'-----------------------------------------------------------
Private Sub cmdNewVersion_Click()
'-----------------------------------------------------------
' REVISIONS
'changed by Mo Morris 23/2/00
' DPH 17/10/2001 - Check for file creation
' DPH 03/05/2002 - Don't write message to message table unless
'           file is placed in published HTML folder
'-----------------------------------------------------------

Dim oExchange As New clsExchange
Dim bTest As Boolean
Dim sMsg As String
Dim sCabFileName As String

On Error GoTo ErrHandler
    If lvwTrials.SelectedItem Is Nothing Then
        MsgBox "Please select a study"
        Exit Sub
    End If
    
    If lvwTrials.SelectedItem.SubItems(1) = StatusDescription(eTrialStatus.InPreparation) Then   'In preparation
        sMsg = "This study has not been opened.  Do you want to distribute a test version ?"
        If MsgBox(sMsg, vbYesNoCancel) = vbYes Then
            bTest = True
        Else
            Exit Sub
        End If
    Else
        bTest = False
    End If
    
    ' DPH 07/08/2002 - Call new form to Label a study version
    Call frmStudyLabel.InitialiseMe(ClinicalTrialId, msSelectedClinicalTrialName, gnCurrentVersionId(ClinicalTrialId))
    FormCentre frmStudyLabel, Me
    frmStudyLabel.Show vbModal
    
    ' DPH 06/01/2003 Refresh TrialList & focus
    Call RefreshTrialList
    lvwTrials.SetFocus
    
'    Screen.MousePointer = vbHourglass
'
'    Set oExchange = New clsExchange
'
'    '24/2/00, Note that ExportSDD now returns the name of the ceated Cab file (excluding path)
'    sCabFileName = oExchange.ExportSDD(ClinicalTrialId, msSelectedClinicalTrialName, gnCurrentVersionId(ClinicalTrialId))
'
'    ' DPH 17/10/2001 - Make sure cab file has been created
'    If sCabFileName <> "" Then
'        ' Do not write message to message table unless successfully copied
'
'        'Mo Morris 24/2/00
'        Do Until FileExists(gsOUT_FOLDER_LOCATION & sCabFileName)
'            DoEvents
'        Loop
'
'        On Error GoTo CopyFileErr
'
'        'Mo Morris 24/2/00
'        Call FileCopy(gsOUT_FOLDER_LOCATION & sCabFileName, _
'                        gsHTML_FORMS_LOCATION & sCabFileName)
'
'        On Error GoTo ErrHandler
'
'        'Mo Morris 23/2/00 pass sCabFileName to CreateMessageForAllTrialSites
'        Call CreateMessageForAllTrialSites(ClinicalTrialId, _
'                                            msSelectedClinicalTrialName, _
'                                            ExchangeMessageType.NewVersion, sCabFileName)
'
'    Else
'        Call DialogError("Distribute new version aborted - unable to create export file", "Study Definition Distribution")
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'
'    Screen.MousePointer = vbDefault
    
    'Mo Morris 12/12/00
'    Call DialogInformation(msSelectedClinicalTrialName & " has successfully been distributed", "Study Definition Distribution")
    
Exit Sub
CopyFileErr:
    ' If an error copying file
    Screen.MousePointer = vbDefault
    
    Call DialogError("Distribute new version aborted - Error copying file to published HTML folder", "Study Definition Distribution")

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "cmdNewVersion_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'-----------------------------------------------------------
Private Sub CmdOpenTrial_Click()
'-----------------------------------------------------------

    On Error GoTo ErrHandler

    Call UpdateTrialStatus(ClinicalTrialId, msSelectedClinicalTrialName, _
                            eTrialStatus.TrialOpen)
    Call ChangeStatus(eTrialStatus.TrialOpen)       'Converts StatusId to its text description
    Call RefreshTrialList
        
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "CmdOpenTrial_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub

'-----------------------------------------------------------
Private Sub cmdRefresh_Click()
'-----------------------------------------------------------

    Call RefreshTrialList
    
End Sub

'-----------------------------------------------------------
Private Sub cmdKeywordSearch_Click()
'-----------------------------------------------------------
' Iterate through list items collection of the list view control
' and test if the text in the search text box is contained in the
' keyword subitem.  Remove list items that do not contain the
' search text.
'-----------------------------------------------------------
' REVISIONS
' DPH 29/09/2006 - Bug 2757 - corrected Keywords appearing under wrong column header
'-----------------------------------------------------------

Dim oTrialItem As ListItem
Dim rsTrialList As ADODB.Recordset
Dim skey As String
Dim sSearchWord As String
Dim sSQL As String
Dim sRecruitmentSQL As String
Dim rsTotal As ADODB.Recordset

    On Error GoTo ErrHandler
    'Prepare searchWord for use by the Like Operator.
    'The "%" at either end of the user entered search word stands for
    'any number of wildcard characters
    
    ' PN 15/09/99 changed wildcard char to be ansi standard as used with ado
    sSearchWord = "%" & txtKeywordSearch.Text & "%"
    
    Me.MousePointer = vbHourglass
    
    'Disable command buttons
    DisableCommands
    
    'Remove existing items
    lvwTrials.ListItems.Clear
    
    ' Get the list of trials, excluding the library
    Set rsTrialList = New ADODB.Recordset
    sSQL = "SELECT * FROM ClinicalTrial WHERE ClinicalTrialId > 0 " _
        & " AND Keywords Like '" & sSearchWord & "'"
    rsTrialList.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    ' While the record is not the last record, add a ListItem object.
    Do While Not rsTrialList.EOF
    
        ' Store the sponsor and trial id in the listitem key
        skey = msKEY_SEPARATOR & rsTrialList![ClinicalTrialId]
        
        ' DPH 29/09/2006 - Bug 2757 - Keywords appearing under wrong column header
        ' collect recruitment numbers
        sRecruitmentSQL = "SELECT COUNT(*)"
        sRecruitmentSQL = sRecruitmentSQL & " FROM TrialSubject"
        sRecruitmentSQL = sRecruitmentSQL & " WHERE TrialSubject.ClinicalTrialID=" & rsTrialList![ClinicalTrialId]
        
        Set rsTotal = New ADODB.Recordset
        rsTotal.Open sRecruitmentSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
        
        ' add the listitem
        Set oTrialItem = lvwTrials.ListItems.Add(, skey, _
            rsTrialList![ClinicalTrialName], gsTRIAL_LABEL, gsTRIAL_LABEL)
        
        With oTrialItem
            ' look up the description of the trial status
            .SubItems(1) = gsTrialStatus(rsTrialList![statusId])
            ' look up the description of the trial phase
            .SubItems(2) = gsTrialPhase(rsTrialList![PhaseId])
            .SubItems(3) = rsTrialList![ExpectedRecruitment]
            ' DPH 29/09/2006 - Bug 2757 - Keywords appearing under wrong column header
            ' Add recruitment total + version number
            .SubItems(4) = rsTotal.Fields(0).Value
            .SubItems(5) = GetLatestVersionOfTrialAvailable(rsTrialList![ClinicalTrialId])
            .SubItems(6) = rsTrialList![keywords]
        End With
        
        rsTotal.Close
        
        rsTrialList.MoveNext   ' Move to next record.
    Loop
    
    ' Close the recordset
    rsTrialList.Close
    Set rsTrialList = Nothing
    Set rsTotal = Nothing
    
    Me.MousePointer = vbDefault
    
    'make sure that no trial is selected initially
    Set lvwTrials.SelectedItem = Nothing
        
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "cmdKeywordSearch_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub

'-----------------------------------------------------------
Private Sub CmdSuspend_Click()
'-----------------------------------------------------------
    
    Call UpdateTrialStatus(ClinicalTrialId, msSelectedClinicalTrialName, _
                            eTrialStatus.Suspended)
    Call ChangeStatus(eTrialStatus.Suspended)
    Call RefreshTrialList

End Sub

'-----------------------------------------------------------
Private Sub CmdTrialSites_Click()
'-----------------------------------------------------------
' REVISIONS
' DPH 20/08/2002 - Use frmTrialSiteAdminVersioning form
'-----------------------------------------------------------

    On Error GoTo ErrHandler
    
    '   TA 16/10/2000: Select Study selected in listview when showing TrialSite admin
    If lvwTrials.SelectedItem Is Nothing Then 'SDM 01/02/00 SR2861
        ' DPH 20/08/2002 - Use frmTrialSiteAdminVersioning form
'        frmTrialSiteAdmin.Display (DisplaySitesByTrial)
        Call frmTrialSiteAdminVersioning.Display(msDatabase, DisplaySitesByTrial)
    Else
        ' DPH 20/08/2002 - Use frmTrialSiteAdminVersioning form
'        frmTrialSiteAdmin.Display DisplaySitesByTrial, lvwTrials.SelectedItem.Text
        Call frmTrialSiteAdminVersioning.Display(msDatabase, DisplaySitesByTrial, lvwTrials.SelectedItem.Text)
    End If
    
        
    Exit Sub
    
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "CmdTrialSites_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub

'-----------------------------------------------------------
Private Sub Form_Load()
'-----------------------------------------------------------
' REVISIONS
' DPH 03/09/2002 - Added Latest Version Number
'-----------------------------------------------------------

    On Error GoTo ErrHandler
    
    ' Turn on key preview for form, so that F1 (Help) can be trapped by form
    Me.KeyPreview = True
    Set Me.Icon = frmMenu.Icon
    'Position the controls
    
    ' Initialise icon images from the resource file
    imglistLargeIcons.ListImages.Add , gsTRIAL_LABEL, LoadResPicture(gsTRIAL_LABEL, vbResIcon)
    imglistSmallIcons.ListImages.Add , gsTRIAL_LABEL, LoadResPicture(gsTRIAL_LABEL, vbResIcon)
    
    ' Create an object variable for the ColumnHeader object.
    ' Add ColumnHeaders with appropriate widths
    With lvwTrials

        
        .ColumnHeaders.Add , , "Name", 1700
        .ColumnHeaders.Add , , "Status", 1000
        .ColumnHeaders.Add , , "Phase", 500
        .ColumnHeaders.Add , , "Recruitment", 800
        .ColumnHeaders.Add , , "Actual Recruitment", 1400
        .ColumnHeaders.Add , , "Latest Version", 800
        .ColumnHeaders.Add , , "Keywords", 15200
        
        .View = lvwReport ' Set View property to report
        
        ' Sort on first column (trial name) ascending
        .SortKey = 0
        .SortOrder = lvwAscending
        .Sorted = True
    
        .Icons = imglistLargeIcons
        .SmallIcons = imglistSmallIcons
    End With
    
'    ' Populate the list
'    Call RefreshTrialList
    
    'Mo Morris 6/3/00, SR2764
    'Command Button disabling placed into DisableCommands which is now called
    'from RefreshTrialList and cmdKeywordSearch_Click

    
    ' NCJ 14/12/99
    'changed Mo Morris 13/1/00, EnableUsersButtons now called from lvwTrials_ItemClick
    'Call EnableUsersButtons
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "Form_Load")
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
Private Sub EnableUsersButtons()
'---------------------------------------------------------------------
' NCJ 14/12/99
' Enable/disable buttons according to current user's access rights
'---------------------------------------------------------------------

    If goUser.CheckPermission(gsFnDistribNewVersionOfStudyDef) And Not (lvwTrials.SelectedItem Is Nothing) Then
        cmdNewVersion.Enabled = True
    Else
        cmdNewVersion.Enabled = False
    End If
    
    If goUser.CheckPermission(gsFnChangeTrialStatus) And Not (lvwTrials.SelectedItem Is Nothing) Then
        cmdInPreparation.Enabled = (lvwTrials.SelectedItem.SubItems(1) = "Open")
        CmdOpenTrial.Enabled = True
        CmdSuspend.Enabled = True
        CmdCloseRecruitment.Enabled = True
        CmdCloseFollowUp.Enabled = True
    Else
        cmdInPreparation.Enabled = False
        CmdOpenTrial.Enabled = False
        CmdSuspend.Enabled = False
        CmdCloseRecruitment.Enabled = False
        CmdCloseFollowUp.Enabled = False
    End If
    

End Sub

'-----------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'-----------------------------------------------------------
On Error GoTo ErrHandler

    If KeyCode = vbKeyF1 Then               ' Show user guide
        'ShowDocument Me.hWnd, gsMACROUserGuidePath
        
        'REM 07/12/01 - New Call to MACRO Help
        Call MACROHelp(Me.hWnd, App.Title)
        
    End If
        
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "Form_KeyDown")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub

'-----------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'-----------------------------------------------------------
    
    mnSelectedClinicalTrialId = 0
    msSelectedClinicalTrialName = ""
            
End Sub

'-----------------------------------------------------------
Private Sub lvwTrials_ColumnClick(ByVal ColumnHeader As ColumnHeader)
'-----------------------------------------------------------
On Error GoTo ErrHandler

    ' When a ColumnHeader object is clicked, the ListView control is
    ' sorted by the subitems of that column.
    ' Set the SortKey to the Index of the ColumnHeader - 1
    lvwTrials.SortKey = ColumnHeader.Index - 1
    
    ' Reverse the sort order
    If lvwTrials.SortOrder = lvwAscending Then
        lvwTrials.SortOrder = lvwDescending
    Else
        lvwTrials.SortOrder = lvwAscending
    End If
    
    ' Set Sorted to True to sort the list.
    lvwTrials.Sorted = True
        
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "lvwTrials_ColumnClick")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub

'-----------------------------------------------------------
Private Sub lvwTrials_ItemClick(ByVal Item As ListItem)
'-----------------------------------------------------------
On Error GoTo ErrHandler

    mnSelectedClinicalTrialId = Mid(Item.Key, InStr(Item.Key, msKEY_SEPARATOR) + 1)
    msSelectedClinicalTrialName = lvwTrials.SelectedItem
    
    'WillC 24/5/00   SR3492
    msSelectedTriaStatus = lvwTrials.SelectedItem.SubItems(1)
    
    
    '  Enable options once trial has been selected
    'changed Mo Morris 13/1/00 SR2662
    'fraOptions.Enabled = True
    cmdNewVersion.Enabled = True
    CmdHistory.Enabled = True
    CmdTrialSites.Enabled = True
    Call EnableUsersButtons
        
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "lvwTrials_ItemClick")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub

'-----------------------------------------------------------
Private Sub txtKeywordSearch_Change()
'-----------------------------------------------------------
On Error GoTo ErrHandler

    If txtKeywordSearch.Text <> "" Then
        cmdKeywordSearch.Enabled = True
    Else
        cmdKeywordSearch.Enabled = False
    End If
        
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "txtKeywordSearch_Change")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub

'-----------------------------------------------------------
Private Sub ChangeStatus(statusId As Integer)
'-----------------------------------------------------------
' changes the state of the buttons
'-----------------------------------------------------------

    'cboStatus.ListIndex = StatusId
    
    'Select Case gsTrialStatus(statusId)
    '    Case "In preparation"
    '        CmdOpenTrial.Enabled = True
    '        CmdCloseRecruitment.Enabled = True
    '        CmdCloseFollowUp.Enabled = True
    '        CmdSuspend.Enabled = False
    '    Case "Open"
    '        CmdOpenTrial.Enabled = False
    '        CmdCloseRecruitment.Enabled = True
    '        CmdCloseFollowUp.Enabled = True
    '        CmdSuspend.Enabled = True
    '    Case "Closed to recruitment"
    '        CmdOpenTrial.Enabled = False
    '        CmdCloseRecruitment.Enabled = False
    '        CmdCloseFollowUp.Enabled = True
    '        CmdSuspend.Enabled = True
    '    Case "Closed to follow up"
    '        CmdOpenTrial.Enabled = False
    '        CmdCloseRecruitment.Enabled = False
    '        CmdCloseFollowUp.Enabled = False
    '        CmdSuspend.Enabled = False
    '        'cmdNewVersion.Enabled = False
    '    Case "Suspended"                    'Grid removed. MsgBox shows history instead
    '        CmdSuspend.Enabled = False
    '        'grdVersions.Row = grdVersions.Rows - 2
    '        'grdVersions.Col = 1
    '        'Select Case grdVersions.Text
    '             'Case "Open"
    '                'CmdOpenTrial.Enabled = True
    '                'CmdCloseRecruitment.Enabled = True
    '                'CmdCloseFollowUp.Enabled = True
    '            'Case "Closed to recruitment"
    '                'CmdOpenTrial.Enabled = False
    '                'CmdCloseRecruitment.Enabled = True
    '                'CmdCloseFollowUp.Enabled = True
    '            'Case "Closed to follow up"
    '                'CmdOpenTrial.Enabled = False
    '                'CmdCloseRecruitment.Enabled = False
    '                'CmdCloseFollowUp.Enabled = True
    '        'End Select
    'End Select

End Sub

'--------------------------------------------------------------------
Private Sub UpdateTrialStatus(vClinicalTrialId As Long, _
                                vClinicalTrialName As String, _
                                vStatusId As Integer)
'--------------------------------------------------------------------
    
  On Error GoTo ErrHandler
  Screen.MousePointer = vbHourglass
   
   
   
   
   
    'Begin transaction
    TransBegin
    
    Call gdsUpdateTrialStatus(vClinicalTrialId, _
                                gnCurrentVersionId(vClinicalTrialId), _
                                vStatusId)
    
    Call CreateMessageForAllTrialSites(vClinicalTrialId, _
                                        vClinicalTrialName, _
                                        vStatusId)
    'End transaction
    TransCommit
    
    Screen.MousePointer = vbDefault
        
Exit Sub
ErrHandler:

    'RollBack transaction
    TransRollBack

    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "UpdateTrialStatus")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub

'--------------------------------------------------------------------
Private Sub CreateMessageForAllTrialSites(vClinicalTrialId As Long, _
                                            vClinicalTrialName As String, _
                                            vStatusId As Integer, Optional sCabFileName As String)
'--------------------------------------------------------------------
'changed Mo Morris 23/2/00  sCabFileName added as a Parameter
' DPH 29/05/2002 - Changed SQL to disallow inactive sites from distribution
'--------------------------------------------------------------------
Dim rsTrialSites As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler
    
    ' DPH 29/05/2002 - Changed SQL to disallow inactive sites from distribution
    'sSQL = "SELECT TrialSite from TrialSite WHERE ClinicalTrialId = " & vClinicalTrialId
    sSQL = "SELECT TrialSite.TrialSite From TrialSite, Site " & _
        "WHERE (TrialSite.ClinicalTrialId = " & vClinicalTrialId & ") AND " & _
        " (TrialSite.TrialSite = Site.Site) AND (Site.SiteStatus = 0)"
    
    Set rsTrialSites = New ADODB.Recordset
    'changed by Mo Morris 4/1/00, cursor type changed from adOpenForwardOnly to adOpenStatic
    rsTrialSites.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adOpenStatic, adCmdText
    
    Do While Not rsTrialSites.EOF
    
        '   ATN 18/2/99
        '   MSMQ prefix removed because MSMQ no longer used.
        Call CreateStatusMessage(vStatusId, vClinicalTrialId, _
                                  vClinicalTrialName, rsTrialSites!TrialSite, sCabFileName)
            
        rsTrialSites.MoveNext
    Loop
        
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "CreateMessageForAllTrialSites")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    

End Sub

'--------------------------------------------------------------------
Private Sub DisableCommands()
'--------------------------------------------------------------------
'Sub added Mo Morris 6/3/00, SR2764
'--------------------------------------------------------------------

    'Disable options until trial has been selected
    'changed Mo Morris 13/1/00 SR2662
    'fraOptions.Enabled = True
    cmdNewVersion.Enabled = False
    cmdInPreparation.Enabled = False
    CmdOpenTrial.Enabled = False
    CmdSuspend.Enabled = False
    CmdCloseRecruitment.Enabled = False
    CmdCloseFollowUp.Enabled = False
    CmdHistory.Enabled = False
    CmdTrialSites.Enabled = False

End Sub

'--------------------------------------------------------------------
Private Sub CreateNewVersion()
'--------------------------------------------------------------------
'
'--------------------------------------------------------------------

End Sub

'--------------------------------------------------------------------------------
Private Function GetLatestVersionOfTrialAvailable(lTrialId As Long) As Integer
'--------------------------------------------------------------------------------
' Extract
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim rsVersion As ADODB.Recordset
Dim nVersion As Integer

    On Error GoTo ErrorHandler
    
    sSQL = "SELECT Max(StudyVersion) AS MaxVersion FROM StudyVersion WHERE ClinicalTrialId = " & lTrialId
    
    Set rsVersion = New ADODB.Recordset
    rsVersion.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

    If Not rsVersion.EOF Then
        If IsNull(rsVersion("MaxVersion")) Then
            nVersion = 0
        Else
            nVersion = rsVersion("MaxVersion")
        End If
    Else
        nVersion = 0
    End If
    rsVersion.Close
    Set rsVersion = Nothing
    
    GetLatestVersionOfTrialAvailable = nVersion
    
Exit Function
ErrorHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "GetLatestVersionOfTrialAvailable")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
  End Select
End Function

'---------------------------------------------------------------------------------
Public Sub Display(Optional sDatabase As String)
'---------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------
    
    ' Populate the list
    Load Me
    FormCentre Me, frmMenu
    Call RefreshTrialList
    msDatabase = sDatabase
    Me.Caption = "List of Studies " & "[" & goUser.DatabaseCode & "]"
    Me.Show vbModal

End Sub
