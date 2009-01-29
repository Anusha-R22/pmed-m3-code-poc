VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTrialList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of Studies"
   ClientHeight    =   7245
   ClientLeft      =   4545
   ClientTop       =   2640
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7245
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraAccess 
      Caption         =   "Study Access Mode"
      Height          =   615
      Left            =   720
      TabIndex        =   8
      Top             =   6480
      Width           =   5655
      Begin VB.OptionButton optFC 
         Caption         =   "Full control"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optRO 
         Caption         =   "Read only"
         Height          =   255
         Left            =   4320
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optLO 
         Caption         =   "Layout only"
         Height          =   255
         Left            =   3000
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optRW 
         Caption         =   "Read write"
         Height          =   195
         Left            =   1680
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6900
      TabIndex        =   6
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   8220
      TabIndex        =   7
      Top             =   6480
      Width           =   1215
   End
   Begin VB.PictureBox picSearch 
      Height          =   540
      Left            =   60
      ScaleHeight     =   480
      ScaleWidth      =   9315
      TabIndex        =   1
      Top             =   108
      Width           =   9375
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   330
         Left            =   6000
         TabIndex        =   5
         Top             =   108
         Width           =   1065
      End
      Begin VB.CommandButton cmdKeywordSearch 
         Caption         =   "&Search"
         Height          =   330
         Left            =   4725
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
      Height          =   5640
      Left            =   60
      TabIndex        =   0
      Top             =   735
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   9948
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
   Begin VB.Menu mnuV 
      Caption         =   "&View"
      Begin VB.Menu mnuVIcon 
         Caption         =   "Lar&ge icons"
      End
      Begin VB.Menu mnuVSmallIcon 
         Caption         =   "S&mall icons"
      End
      Begin VB.Menu mnuVList 
         Caption         =   "&List"
      End
      Begin VB.Menu mnuVReport 
         Caption         =   "&Details"
      End
   End
End
Attribute VB_Name = "frmTrialList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998 - 2006. All Rights Reserved
'   File:       frmTrialList.frm
'   Author:     Andrew Newbigging June 1997
'   Purpose:    Displays list of trials and allows user to open one for editing.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'  Andrew Newbigging et al., 4/06/97 - 20/09/99, various changes
'   MM  20/09/99    Removed DeleteTrial() and  menu options mnuFCopy_Click(),
'                   mnuFClose_Click(), mnuFDelete_Click(), mnuFOpenCopy_Click(),
'                   mnuFOpenTrial_Click(), mnuFNewTrial_Click()
'   PN  30/09/99    Amended RefreshTrialList() to specify full SQL statment
'   WillC 11/11/99  Added the error handlers
'   Mo Morris   15/11/99    DAO to ADO conversion
'   NCJ 7 Dec 99, SR 2217 - Use "%" as wildcard character rather than "?" in Keyword search
'       SR 2162 - Process keypress on KeyDown rather than KeyUp
'   NCJ 13 Dec 99 - Changed Ids to Long
'   Mo Morris   4/1/00
'   gbNoSelectionMade added for assessing wether a trial has been selected prior
'   to opening it with a double click (SR 1754)
'   WillC 22/2/200 SR2886 Changed the height of lvwTrials to accomodate the buttons
'   TA 29/04/2000   subclassing removed
'   Mo  15/11/2000  ReplaceCharacters call replaced by call to VB Function Replace
'   MLM 14/06/02:   Check that the selected study has not been deleted from db.
'   MLM 18/06/02:   When refreshing the list, pre-select the 1st item in the list.
'   NCJ 4 May 06 - Allow user to set Study Access Mode
'   NCJ 8 Jun 06 - Added Full Control to SAM
'   NCJ 12 Jul 06 - Only allow MU if relevant switch is set
'   NCJ 26 Sept 06 - Corrected text of Read write option button
'   NCJ 24 Oct 06 - Default access mode is RO
'------------------------------------------------------------------------------------'
'-----------------------------------------------------------
' FORM: frmTrialList
'
' Displays a list of all trials defined in the system,
' and allows the user to choose one, for editing, copying
' or deletion, or to create a new one.  A list can be
' printed (using the report 'TrialList').
' The form has 2 'modes', one to allow the user to choose
' a study definition, for editing, copying
' or deletion, or to create a new one.  The other mode
' is used to allow the user to open a read-only copy
' of a study definition
'-----------------------------------------------------------

Private msMode As String    'stores the value of the Mode property
'Private Const msUPDATE_MODE = "Study definition"
'Private Const msCOPY_MODE = "Copy other"
Private Const msKEY_SEPARATOR = "_"

Private mlSelectedClinicalTrialId As Long
Private msSelectedClinicalTrialName As String
'code added by Mo Morris 4/1/00
Private gbNoSelectionMade As Boolean

Private meAccessMode As eSDAccessMode     ' NCJ 9 May 06

'---------------------------------------------------------------------
Private Sub cmdExit_Click()
'---------------------------------------------------------------------
'Close the form
'---------------------------------------------------------------------
    
    Unload Me
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdOpen_Click()
'---------------------------------------------------------------------
' Open a trial if one is chosen using  the listview double click
'---------------------------------------------------------------------

    Call lvwTrials_DblClick

End Sub

'---------------------------------------------------------------------
Private Sub lvwTrials_KeyDown(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------
' Interpret RETURN as double click
' NCJ 7 Dec 99 - Moved here from KeyUp (SR 2162)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    Select Case KeyCode
    Case vbKeyReturn
        Call lvwTrials_DblClick
    Case vbKeyInsert
        Call NewTrial
        
    ' MM 20/09/99
    ' removed DeleteTrial call
    ' trials can not be deleted here
    
    Case vbKeyEscape
        mlSelectedClinicalTrialId = 0
        msSelectedClinicalTrialName = ""

        Me.Hide
    End Select
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "lvwTrials_KeyDown")
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
Public Property Get SelectedClinicalTrialId() As Long
'---------------------------------------------------------------------

    SelectedClinicalTrialId = mlSelectedClinicalTrialId

End Property

'---------------------------------------------------------------------
Public Property Get SelectedClinicalTrialName() As String
'---------------------------------------------------------------------

    SelectedClinicalTrialName = msSelectedClinicalTrialName

End Property

'---------------------------------------------------------------------
Public Property Get Mode() As String
'---------------------------------------------------------------------
    
    Mode = msMode

End Property

'---------------------------------------------------------------------
Public Property Let Mode(ByVal sMode As String)
'---------------------------------------------------------------------
' Whether we just want to select a study (gsREAD) or open it for editing (gsUPDATE)
' NCJ 9 May 06 - Invisiblise access modes if we only want to read study def
' NCJ 12 Jul 06 - Check Multi-User setting too
'---------------------------------------------------------------------

    msMode = sMode
    fraAccess.Visible = (msMode = gsUPDATE And frmMenu.AllowMU)
    If msMode = gsREAD Then
        optRO.Value = True
    End If
    
End Property

'---------------------------------------------------------------------
Public Property Get StudyAccessMode() As eSDAccessMode
'---------------------------------------------------------------------
' NCJ 10 May 06 - Study Access Mode (only relevant when opening a study for editing)
'---------------------------------------------------------------------

    StudyAccessMode = meAccessMode

End Property

'---------------------------------------------------------------------
Private Sub RefreshTrialList()
'---------------------------------------------------------------------
' SUB: RefreshTrialList
'
' Reads the list of trials from the database and displays
' it in the form.
'
' MLM 18/06/02: Select first item in the list.
'---------------------------------------------------------------------

Dim itmTrial As ListItem
Dim rsTrialList As ADODB.Recordset
Dim msKey As String
Dim oTrialList As clsTextList
Dim sSQL As String

On Error GoTo ErrHandler

    Me.MousePointer = vbHourglass
    
    'Remove existing items
    lvwTrials.ListItems.Clear
    'MLM 18/06/02
    gbNoSelectionMade = True
    
    ' Get the list of trials, excluding the library
    sSQL = "SELECT * FROM ClinicalTrial WHERE ClinicalTrialId > 0 ORDER by ClinicalTrialName"
    Set rsTrialList = New ADODB.Recordset
    rsTrialList.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    ' PN 09/09/99 - read the available trial types
    ' load the possible trial types with a blank value
    Set oTrialList = New clsTextList
    ' PN 30/09/99 - change to specify full SQL statement for ODBC compliance with ADO
    oTrialList.Load "Select * from TrialType", True, MacroADODBConnection
        
    ' While the record is not the last record, add a ListItem object.
    With rsTrialList
        While Not .EOF
        
            ' Store the sponsor and trial id in the listitem key
            msKey = msKEY_SEPARATOR & rsTrialList![ClinicalTrialId]
                    
            ' add the listitem
            Set itmTrial = lvwTrials.ListItems.Add(, msKey, _
                rsTrialList![ClinicalTrialName], gsTRIAL_LABEL, gsTRIAL_LABEL)
               
            ' look up the description of the trial status
            itmTrial.SubItems(1) = gsTrialStatus(rsTrialList![statusId])
            ' look up the description of the trial phase
            itmTrial.SubItems(2) = gsTrialPhase(rsTrialList![PhaseId])
            itmTrial.SubItems(3) = rsTrialList![ExpectedRecruitment]
    
            ' PN 09/09/99 - read the available trial types from the object
            ' look up the description of the Trial Type
            itmTrial.SubItems(4) = oTrialList.Item(rsTrialList.Fields("TrialTypeId"))
            
            If IsNull(rsTrialList![Keywords]) Then
                itmTrial.SubItems(5) = vbNullString
            Else
                itmTrial.SubItems(5) = rsTrialList![Keywords]
            End If
            
            'MLM 18/06/02: "Select" the 1st item in the list
            If gbNoSelectionMade Then
                Call lvwTrials_ItemClick(itmTrial)
            End If
            
            .MoveNext   ' Move to next record.
        Wend
        
        ' Close the recordset
        .Close
    End With
    Set rsTrialList = Nothing
    
    ' PN 09/09/99
    ' destroy the trial list object
    Set oTrialList = Nothing
    
    lvw_SetAllColWidths lvwTrials, LVSCW_AUTOSIZE_USEHEADER
    
    Me.MousePointer = vbDefault
    
    'make sure that no trial is selected initially
    'Set lvwTrials.SelectedItem = Nothing
    
    'Disable the Search command and clear down the search text
    cmdKeywordSearch.Enabled = False
    txtKeywordSearch.Text = ""
    
    
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
            End
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub NewTrial()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    Screen.MousePointer = vbHourglass
    ' NCJ 10 May 06 - Mode not used any more
'    frmMenu.Mode = gsSTUDY_DEFINITION_MODE
    Me.Hide
    frmMenu.NewTrial
    Screen.MousePointer = vbDefault
    
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "NewTrial")
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
Private Sub CopyTrial()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    MsgBox "This option has been temporarily disabled."
    
    Exit Sub
    
    If Not (lvwTrials.SelectedItem Is Nothing) Then
    
        Screen.MousePointer = vbHourglass
        
        ' Copy the selected trial, passing  ClinicalTrialId
        frmMenu.CopyTrial mlSelectedClinicalTrialId
        Me.Hide
    
    End If
    
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "CopyTrial")
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
Private Sub OpenTrial()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    If lvwTrials.SelectedItem Is Nothing Then
    
        mlSelectedClinicalTrialId = 0
        msSelectedClinicalTrialName = ""
    
    End If
        
    Me.Hide
    'Unload Me
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "OpenTrial")
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
Public Sub OpenCopy()
'---------------------------------------------------------------------
' NCJ 9 May 06 - This seems to be identical to OpenTrial, so ignore!
'---------------------------------------------------------------------

'    On Error GoTo ErrHandler
'
'    If lvwTrials.SelectedItem Is Nothing Then
'
'        mlSelectedClinicalTrialId = 0
'        msSelectedClinicalTrialName = ""
'
'    End If
'
'    Me.Hide
'
'Exit Sub
'ErrHandler:
'    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
'                                    "OpenCopy")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            End
'    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdRefresh_Click()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    Call RefreshTrialList
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdRefresh_Click")
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
Private Sub cmdKeywordSearch_Click()
'---------------------------------------------------------------------
' Iterate through list items collection of the list view control
' and test if the text in the search text box is contained in the
' keyword subitem.  Remove list items that do not contain the
' search text.
' NCJ 7 Dec 99, SR 2217 - Use "%" as wildcard character rather than "?"
' MLM 18/06/02: Select first item in the list
'---------------------------------------------------------------------

Dim itmTrial As ListItem
Dim rsTrialList As ADODB.Recordset
Dim msKey As String
Dim msSearchWord As String
Dim sSQL As String

    On Error GoTo ErrHandler
    
    'Prepare searchWord for use by the Like Operator.
    'The "%" at either end of the user entered search word stands for
    'any number of wildcard characters
    '   ATN 24/2/2000 SR 3099
    '   Need to use 2 single quotes to search successfully for keywors containing quotes.
    msSearchWord = "%" + Replace(Replace(txtKeywordSearch.Text, """", "'"), "'", "''") + "%"
    '   Also need to remove pipe characters, otherwise Access will crash
    msSearchWord = Replace(msSearchWord, "|", "")
    
    Me.MousePointer = vbHourglass
    
    'Remove existing items
    lvwTrials.ListItems.Clear
    'MLM 18/06/02
    gbNoSelectionMade = True
    
    ' Get the list of trials, excluding the library
    sSQL = "SELECT * FROM ClinicalTrial " _
        & " WHERE ClinicalTrialId > 0 " _
        & " AND Keywords Like '" & msSearchWord & "'"
    Set rsTrialList = New ADODB.Recordset
    rsTrialList.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    ' While the record is not the last record, add a ListItem object.
    While Not rsTrialList.EOF
    
        ' Store the sponsor and trial id in the listitem key
        msKey = msKEY_SEPARATOR & rsTrialList![ClinicalTrialId]
                
        ' add the listitem
        Set itmTrial = lvwTrials.ListItems.Add(, msKey, _
            rsTrialList![ClinicalTrialName], gsTRIAL_LABEL, gsTRIAL_LABEL)
            
        ' look up the description of the trial status
        itmTrial.SubItems(1) = gsTrialStatus(rsTrialList![statusId])
        ' look up the description of the trial phase
        itmTrial.SubItems(2) = gsTrialPhase(rsTrialList![PhaseId])
        itmTrial.SubItems(3) = rsTrialList![ExpectedRecruitment]
        itmTrial.SubItems(5) = rsTrialList![Keywords]
        
        'MLM 18/06/02: "Select" the 1st item in the list
        If gbNoSelectionMade Then
            Call lvwTrials_ItemClick(itmTrial)
        End If
        
        rsTrialList.MoveNext   ' Move to next record.
    Wend
    
    ' Close the recordset
    rsTrialList.Close
    Set rsTrialList = Nothing
    
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
            End
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler

    Me.Icon = frmMenu.Icon
    
    ' Turn on key preview for form, so that F1 (Help) can be trapped by form
    Me.KeyPreview = True
    
    'Position the controls
    
    picSearch.Left = 100
    picSearch.Width = Me.Width - 300
    picSearch.Top = 100
    lvwTrials.Left = 100
    lvwTrials.Width = Me.Width - 300
    lvwTrials.Top = picSearch.Top + picSearch.Height + 100
    
   ' lvwTrials.Height = Me.Height - lvwTrials.Top - 800
   ' WillC 22/2/200 SR2886 Changed the height of lvwTrials to accomodate the buttons
   ' We also need to change the listview to a version that supports FullRowSelect
'    lvwTrials.Height = Me.Height - lvwTrials.Top - 1300
 
    fraAccess.Top = Me.ScaleHeight - fraAccess.Height - 100
    cmdOpen.Top = fraAccess.Top + (fraAccess.Height - cmdOpen.Height) / 2
    cmdExit.Top = cmdOpen.Top
    
    lvwTrials.Height = fraAccess.Top - 100 - lvwTrials.Top
    
    ' Create an object variable for the ColumnHeader object.
    Dim clmX As ColumnHeader
    ' Add ColumnHeaders with appropriate widths
    Set clmX = lvwTrials.ColumnHeaders.Add(, , "Name", 1700)
    Set clmX = lvwTrials.ColumnHeaders.Add(, , "Status", 1000)
    Set clmX = lvwTrials.ColumnHeaders.Add(, , "Phase", 500)
    Set clmX = lvwTrials.ColumnHeaders.Add(, , "Recruitment", 800)
    'TA 14/10/2000: changed trial type to study type
    Set clmX = lvwTrials.ColumnHeaders.Add(, , "Study Type", 1700)
    Set clmX = lvwTrials.ColumnHeaders.Add(, , "Keywords", 15200)
    
    
    lvwTrials.View = lvwReport ' Set View property to report
    
    ' Sort on first column (trial name) ascending
    lvwTrials.SortKey = 0
    lvwTrials.SortOrder = lvwAscending
    lvwTrials.Sorted = True
    
    ' Initialise icon images from the resource file
    Dim imgX As ListImage
    Set imgX = imglistLargeIcons.ListImages.Add(, gsTRIAL_LABEL, _
                    LoadResPicture(gsTRIAL_LABEL, vbResIcon))
    Set imgX = imglistSmallIcons.ListImages.Add(, gsTRIAL_LABEL, _
                    LoadResPicture(gsTRIAL_LABEL, vbResIcon))
                    
    lvwTrials.Icons = imglistLargeIcons
    lvwTrials.SmallIcons = imglistSmallIcons
    
    ' Populate the list
    Call RefreshTrialList
    
    ' NCJ 12 Jun 06 - Default to Full Control
    ' NCJ 24 Oct 06 - Default to Read Only (Roche request)
    optRO.Value = True
    
    'code added by Mo Morris 4/1/00
    'MLM 18/06/02: code removed again.. now done in RefreshTrialList
    'gbNoSelectionMade = True
    
 '   FormCentre Me
    
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "Form_Load", Err.Source) = OnErrorAction.Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------

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
            End
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'---------------------------------------------------------------------

    mlSelectedClinicalTrialId = 0
    msSelectedClinicalTrialName = ""
    
End Sub

'---------------------------------------------------------------------
Private Sub lvwTrials_ColumnClick(ByVal ColumnHeader As ColumnHeader)
'---------------------------------------------------------------------

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
            End
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub lvwTrials_DblClick()
'---------------------------------------------------------------------
' MLM 14/06/02: Check that selected study hasn't been deleted.
'---------------------------------------------------------------------

Dim rsStudy As ADODB.Recordset

    On Error GoTo ErrHandler
    
    'code added by Mo Morris 4/1/00
    If gbNoSelectionMade Then
        MsgBox ("Please select a study.")
        Exit Sub
    Else
        'MLM 14/06/02 CBB 2.2.15/11: As a courtesy to the user, check that the study they selected
        'is still in the db before closing this form.
        Set rsStudy = New ADODB.Recordset
        rsStudy.Open "SELECT ClinicalTrialId FROM ClinicalTrial WHERE ClinicalTrialId = " & mlSelectedClinicalTrialId & _
            " AND ClinicalTrialName = '" & msSelectedClinicalTrialName & "'", MacroADODBConnection, _
            adOpenForwardOnly, adLockReadOnly, adCmdText
        If rsStudy.EOF Then
            'study has been deleted
            'Wording makes sense in context of both opening and deleting studies.
            DialogInformation "Another user has deleted the selected study."
            lvwTrials.ListItems.Remove lvwTrials.SelectedItem.Index
            gbNoSelectionMade = True
            Exit Sub
        End If
    End If

' NCJ 9 May 06 - No difference between OpenTrial and OpenCopy!
'    If Mode = msUPDATE_MODE Then
'        OpenTrial
'    ElseIf Mode = msCOPY_MODE Then
'        OpenCopy
'    End If

    Call OpenTrial
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "lvwTrials_DblClick")
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
Private Sub lvwTrials_ItemClick(ByVal Item As ListItem)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    mlSelectedClinicalTrialId = Mid(Item.Key, InStr(Item.Key, msKEY_SEPARATOR) + 1)
    msSelectedClinicalTrialName = lvwTrials.SelectedItem
    
    'code added by Mo Morris 4/1/00
    gbNoSelectionMade = False
    
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
            End
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuVIcon_Click()
'---------------------------------------------------------------------

    lvwTrials.View = lvwIcon
    
End Sub

'---------------------------------------------------------------------
Private Sub mnuVList_Click()
'---------------------------------------------------------------------

    lvwTrials.View = lvwList

End Sub

'---------------------------------------------------------------------
Private Sub mnuVReport_Click()
'---------------------------------------------------------------------

    lvwTrials.View = lvwReport

End Sub

'---------------------------------------------------------------------
Private Sub mnuVSmallIcon_Click()
'---------------------------------------------------------------------

    lvwTrials.View = lvwSmallIcon

End Sub

'---------------------------------------------------------------------
Private Sub optFC_Click()
'---------------------------------------------------------------------

    If optFC.Value = True Then meAccessMode = eSDAccessMode.sdFullControl

End Sub

'---------------------------------------------------------------------
Private Sub optLO_Click()
'---------------------------------------------------------------------

    If optLO.Value = True Then meAccessMode = eSDAccessMode.sdLayoutOnly

End Sub

'---------------------------------------------------------------------
Private Sub optRO_Click()
'---------------------------------------------------------------------

    If optRO.Value = True Then meAccessMode = eSDAccessMode.sdReadOnly
    
End Sub

'---------------------------------------------------------------------
Private Sub optRW_Click()
'---------------------------------------------------------------------
    
    If optRW.Value = True Then meAccessMode = eSDAccessMode.sdReadWrite


End Sub

'---------------------------------------------------------------------
Private Sub txtKeywordSearch_Change()
'---------------------------------------------------------------------

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
            End
    End Select
    
End Sub
