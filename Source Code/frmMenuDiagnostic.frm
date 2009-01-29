VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenu 
   Caption         =   "MACRO 3.0 Diagnostic Utility"
   ClientHeight    =   7530
   ClientLeft      =   3045
   ClientTop       =   4020
   ClientWidth     =   10785
   Icon            =   "frmMenuDiagnostic.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   10785
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   435
      Left            =   360
      TabIndex        =   10
      ToolTipText     =   "Clear all text from the message window"
      Top             =   6600
      Width           =   1215
   End
   Begin VB.ComboBox cboSubject 
      Height          =   315
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   600
      Width           =   975
   End
   Begin VB.ComboBox cboSite 
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   600
      Width           =   1815
   End
   Begin VB.ComboBox cboStudy 
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox txtMsg 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5355
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1080
      Width           =   9735
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   435
      Left            =   8880
      TabIndex        =   0
      ToolTipText     =   "Exit MACRO Diagnostic Utility"
      Top             =   6600
      Width           =   1260
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9720
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrSystemIdleTimeout 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   9720
      Top             =   720
   End
   Begin MSComctlLib.StatusBar sbrMenu 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7155
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Key             =   "UserKey"
            Object.ToolTipText     =   "Name of current user"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Key             =   "RoleKey"
            Object.ToolTipText     =   "Role of current user"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Key             =   "UserDatabase"
            Object.ToolTipText     =   "Name of current Database"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblSubject 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Subject"
      Height          =   255
      Left            =   5520
      TabIndex        =   9
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label Label3 
      Caption         =   "Subject Id:"
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Site:"
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Study:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFSaveCLM 
         Caption         =   "Save CLM File"
      End
      Begin VB.Menu mnuFSavePLM 
         Caption         =   "Save Patient State (PLM)"
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFDataReport 
         Caption         =   "AREZZO Patient Data Listing"
      End
      Begin VB.Menu mnuFTasksReport 
         Caption         =   "AREZZO Tasks Listing"
      End
      Begin VB.Menu mnuTSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFLogOff 
         Caption         =   "Log out"
      End
      Begin VB.Menu mnuFSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuTCLMReport 
         Caption         =   "CLM File Details"
      End
      Begin VB.Menu mnuTSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTCLMMemory 
         Caption         =   "CLM Memory Usage"
      End
      Begin VB.Menu mnuTPLMMemory 
         Caption         =   "PLM Memory Usage"
      End
      Begin VB.Menu mnTS3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTDiagnose 
         Caption         =   "MACRO Patient Data Integrity"
      End
      Begin VB.Menu mnuTSeparator4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTAzConsistency 
         Caption         =   "AREZZO Patient Data Consistency"
      End
      Begin VB.Menu mnuTSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTRecreate 
         Caption         =   "Recreate AREZZO Patient State"
      End
   End
   Begin VB.Menu mnuFHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHUserGuide 
         Caption         =   "&User Guide"
      End
      Begin VB.Menu mnuSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHAboutMacro 
         Caption         =   "&About MACRO"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
' File:         frmMenuDiagnostic.frm
' Copyright:    InferMed Ltd. 2006-2007. All Rights Reserved
' Author:       Nicky Johns, March 2006
' Purpose:      Contains the main form of the MACRO Diagnostic application
'----------------------------------------------------------------------------------------'
'   Revisions:
'   NCJ August 2006 - Added Integrity Check to fix problems arising from Web "Cancel" bug
'   NCJ Jun 2007 - Tidying up ready for general release with MACRO 3.0.77
'   NCJ 20 Jun 07 - Added subject locking around Patient State saving
'----------------------------------------------------------------------------------------'

Option Explicit

Private moArezzo As Arezzo_DM
Private msSelStudyName As String
Private msSelSite As String
Private mlSelStudyId As Long
Private mlSelSubjectId As Long
Private msSubjectSpec As String
Private mbSelSiteRemote As Boolean

'--------------------------------------------------------------------
Private Sub cmdClear_Click()
'--------------------------------------------------------------------
' Clear the message window
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If DialogQuestion("Clear all text from the message window?") = vbYes Then
        txtMsg.Text = ""
        cmdClear.Enabled = False
    End If

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "cmdClear_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Private Sub cmdExit_Click()
'--------------------------------------------------------------------
' Leave the module
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    ' Tidy up
    Call FinaliseMe

    Call ExitMACRO
    Call MACROEnd

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "cmdExit_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Private Sub FinaliseMe()
'--------------------------------------------------------------------
' Do this before we shut down, e.g. when user logs off (but they may log in again)
'--------------------------------------------------------------------
    
    ' Only shut down the ALM if it has been started
    If Not moArezzo Is Nothing Then
        moArezzo.Finish
        Set moArezzo = Nothing
    End If
        
    txtMsg.Text = ""

End Sub

'--------------------------------------------------------------------
Public Sub InitialiseMe()
'--------------------------------------------------------------------
' Initialisations
'--------------------------------------------------------------------
Dim oArezzoMemory As clsAREZZOMemory
Dim sR As String

    On Error GoTo ErrHandler
    
    ' Check for correct user permissions!
    If Not goUser.CheckPermission(gsFnSystemManagement) Then
        DialogInformation "You do not have permission to access the MACRO Diagnostic Utility"
        Call cmdExit_Click
        Exit Sub
    End If
    
    'The following Doevents prevents command buttons ghosting during form load
    DoEvents
     
    ' Tidy up first in case we're re-logging in
    Call FinaliseMe
    
    'Create and initialise a new Arezzo instance
    Set moArezzo = New Arezzo_DM
    
    ' NCJ 29 Jan 03 - Get prolog switches from new ArezzoMemory class
    Set oArezzoMemory = New clsAREZZOMemory
    Call oArezzoMemory.Load(0, goUser.CurrentDBConString)
    'Get the Prolog memory settings using GetPrologSwitches
    Call moArezzo.Init(gsTEMP_PATH, oArezzoMemory.AREZZOSwitches)
    Set oArezzoMemory = Nothing

    ' Load our Diagnostics add-on
    Call moArezzo.ALM.GetPrologResult("ensure_loaded( 'MACRODiagnostics.pc' ), write( '0000' ). ", sR)
    
    If LoadStudies Then
        ' Force initial selections
        Call cboStudy_Click
        Call cboSite_Click
        Call cboSubject_Click
    End If

    Call ShowMessage("MACRO Diagnostic Utility - " & Format(Now, "dd mmm yyyy hh:mm:ss"))
    
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "InitialiseMe", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Public Sub CheckUserRights()
'--------------------------------------------------------------------
' Dummy routine which gets called during MACRO initialisation
'--------------------------------------------------------------------

End Sub

'--------------------------------------------------------------------
Private Sub Form_Load()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    FormCentre Me

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "Form_Load", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Private Sub Form_Resize()
'--------------------------------------------------------------------
' Resize the things on the form
'--------------------------------------------------------------------
Const sglSPACE As Single = 90
Dim sglMinHt As Single

    On Error GoTo ErrHandler
    
    'check that form has not been minimised
    If Me.WindowState = vbMinimized Then Exit Sub
    
    ' Don't allow silly height resize
    ' Say min height of txtMsg is 2 * lblSubject.Height
    sglMinHt = txtMsg.Top + 2 * lblSubject.Height + cmdClear.Height + 2 * sglSPACE + sbrMenu.Height
    If Me.ScaleHeight < sglMinHt Then
        ' Add on the bit for the menus etc.
        Me.Height = sglMinHt + (Me.Height - Me.ScaleHeight)
    End If
    
    ' Position the buttons
    cmdClear.Top = Me.ScaleHeight - sbrMenu.Height - sglSPACE - cmdClear.Height
    cmdExit.Top = cmdClear.Top
    ' Size the message window appropriately
    txtMsg.Height = cmdClear.Top - txtMsg.Top - sglSPACE
    
    ' Don't allow resize beyond left of subject label
    If Me.Width < lblSubject.Left Then
        Me.Width = lblSubject.Left
    End If
    txtMsg.Width = Me.Width - 2 * txtMsg.Left
    cmdExit.Left = txtMsg.Left + txtMsg.Width - cmdExit.Width
    ' Keep a "minimum" string width in subject label
    If txtMsg.Left + txtMsg.Width - lblSubject.Left > Me.TextWidth("WWWWWW") Then
        lblSubject.Width = txtMsg.Left + txtMsg.Width - lblSubject.Left
    End If

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "Form_Resize", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    ' Only shut down the ALM if it has been started
    If Not moArezzo Is Nothing Then
        moArezzo.Finish
        Set moArezzo = Nothing
    End If
    
    Call ExitMACRO
    Call MACROEnd

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "Form_Unload", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Private Sub mnuFExit_Click()
'--------------------------------------------------------------------

    Call cmdExit_Click

End Sub

'--------------------------------------------------------------------
Private Sub mnuFLogOff_Click()
'--------------------------------------------------------------------
' Allow them to log out and log in again
'--------------------------------------------------------------------

    Call UserLogOff(False)
    
End Sub

'--------------------------------------------------------------------
Private Sub mnuFSaveCLM_Click()
'--------------------------------------------------------------------
' Save the CLM file
'--------------------------------------------------------------------

    If msSelStudyName > "" Then
        ShowTime
        Call ShowMessage(DiagSaveCLMFile(msSelStudyName))
    End If

End Sub

'--------------------------------------------------------------------
Private Sub mnuFSavePLM_Click()
'--------------------------------------------------------------------
' Save the PLM file
'--------------------------------------------------------------------
    
    If mlSelStudyId > 0 Then
        ShowTime
        Call ShowMessage(DiagSavePLMFile(msSelStudyName, mlSelStudyId, msSelSite, mlSelSubjectId))
    End If

End Sub

'--------------------------------------------------------------------
Private Sub ShowTime()
'--------------------------------------------------------------------
' Output current time stamp
'--------------------------------------------------------------------

    Call ShowMessage(vbCrLf & Format(Now, "dd/mm/yyyy hh:mm:ss"))

End Sub

'--------------------------------------------------------------------
Private Sub ShowMessage(sText As String)
'--------------------------------------------------------------------
' Add some text to the message window
' AND log it to the Log file
'--------------------------------------------------------------------

    txtMsg.Text = txtMsg.Text & vbCrLf & sText
    ' Move insertion point to the end of the text
    txtMsg.SelStart = Len(txtMsg.Text)
    cmdClear.Enabled = True
    
    Call LogToFile(sText)
    
End Sub

'--------------------------------------------------------------------
Private Sub mnuHAboutMacro_Click()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    frmAbout.Display

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "mnuHAboutMacro_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Private Sub mnuHUserGuide_Click()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    Call MACROHelp(Me.hWnd, App.Title)

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "mnuHUserGuide_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Public Function ForgottenPassword(sSecurityCon As String, sUsername As String, sPassword As String, sErrMsg As String) As Boolean
'---------------------------------------------------------------------
'dummy function for frmNewLogin to compile
'---------------------------------------------------------------------


End Function

'--------------------------------------------
Public Function LoadStudies() As Boolean
'--------------------------------------------
' Populate the Study combo with studies the user has access to
'--------------------------------------------
Dim oStudy As Study
Dim colStudies As Collection

    On Error GoTo ErrLabel
    
    HourglassOn
    
    cboStudy.Clear
    msSelStudyName = ""
    mlSelStudyId = 0
    
    ' Get all studies (which doesn't check the user's Open Subject permission)
    Set colStudies = goUser.GetAllStudies
    
    ' Are there any studies?
    If colStudies.Count = 0 Then
        LoadStudies = False
    Else
        ' Add the studies to the combo
        ' and the study IDs to the ItemData array
        For Each oStudy In colStudies
            cboStudy.AddItem oStudy.StudyName
            cboStudy.ItemData(cboStudy.NewIndex) = oStudy.StudyId
        Next
        cboStudy.ListIndex = 0
        LoadStudies = True
    End If
 
    Set colStudies = Nothing
    Set oStudy = Nothing
    
    HourglassOff

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.LoadStudies"

End Function

'--------------------------------------------
Public Sub LoadSites()
'--------------------------------------------
' Populate the Sites combo with sites the user has access to
' according to the chosen study
'--------------------------------------------
Dim colSites As Collection
Dim oSite As Site

    On Error GoTo ErrLabel
    
    cboSite.Clear
    msSelSite = ""
    HourglassOn
    
    ' Get all sites (which doesn't check the user's Open Subject permission)
    Set colSites = goUser.GetAllSites(cboStudy.ItemData(cboStudy.ListIndex))
    
    ' Are there any sites?
    If colSites.Count > 0 Then
        For Each oSite In colSites
            cboSite.AddItem oSite.Site
            ' Store whether this is a Remote site on a Server
            If (goUser.DBIsServer And oSite.SiteLocation = TypeOfInstallation.RemoteSite) Then
                cboSite.ItemData(cboSite.NewIndex) = TypeOfInstallation.RemoteSite
            Else
                cboSite.ItemData(cboSite.NewIndex) = TypeOfInstallation.Server
            End If
        Next
    End If
    
    If cboSite.ListCount > 0 Then
        cboSite.ListIndex = 0
    Else
        ' Clear out the subject IDs too
        cboSubject.Clear
        mlSelSubjectId = 0
        Call DialogWarning("There are no sites available for this study")
    End If
    
    Set colSites = Nothing
    Set oSite = Nothing
    
    HourglassOff
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.LoadSites"

End Sub

'--------------------------------------------------------------------
Private Sub cboSite_Click()
'--------------------------------------------------------------------
' They clicked on a Site
'--------------------------------------------
    
    If cboSite.ListIndex > -1 Then
        If cboSite.List(cboSite.ListIndex) <> msSelSite Then
            msSelSite = cboSite.List(cboSite.ListIndex)
            ' Is it remote site data on a server?
            mbSelSiteRemote = (cboSite.ItemData(cboSite.ListIndex) = TypeOfInstallation.RemoteSite)
            Call LoadSubjectIds
        End If
    Else
        msSelSite = ""
        mbSelSiteRemote = False
    End If
    Call RememberSubject
    
End Sub

'--------------------------------------------
Private Sub cboStudy_Click()
'--------------------------------------------
' They clicked on a Study
'--------------------------------------------

    ' Any study chosen?
    If cboStudy.ListIndex > -1 Then
        If cboStudy.List(cboStudy.ListIndex) <> msSelStudyName Then
            msSelStudyName = cboStudy.List(cboStudy.ListIndex)
            mlSelStudyId = cboStudy.ItemData(cboStudy.ListIndex)
            Call LoadSites
        End If
    Else
        If msSelStudyName > "" Then
            msSelStudyName = ""
            mlSelStudyId = 0
        End If
    End If
    Call RememberSubject
    
End Sub

'--------------------------------------------
Private Sub RememberSubject()
'--------------------------------------------
' Set up menus according to current study/subject selection
'--------------------------------------------
Dim bSubjEnable As Boolean

    ' Set up the subject spec
    msSubjectSpec = ""
    If msSelStudyName > "" And msSelSite > "" And mlSelSubjectId > 0 Then
        msSubjectSpec = msSelStudyName & "/" & msSelSite & "/" & mlSelSubjectId
    End If
    lblSubject.Caption = msSubjectSpec
    
    ' Enable/disable the CLM menu items
    mnuTCLMReport.Enabled = (msSelStudyName <> "")
    mnuFSaveCLM.Enabled = (msSelStudyName <> "")
    mnuTCLMMemory.Enabled = (msSelStudyName <> "")
    
    ' Now do the subject ones
    bSubjEnable = (msSubjectSpec <> "")
    mnuTAzConsistency.Enabled = bSubjEnable
    mnuTDiagnose.Enabled = bSubjEnable
    mnuFDataReport.Enabled = bSubjEnable
    mnuFTasksReport.Enabled = bSubjEnable
    mnuFSavePLM.Enabled = bSubjEnable
    mnuTPLMMemory.Enabled = bSubjEnable
    ' Reconstructor
    mnuTRecreate.Enabled = bSubjEnable

End Sub

'--------------------------------------------
Private Sub LoadSubjectIds()
'--------------------------------------------
' Load subject ids into combo
'--------------------------------------------
Dim sSQL As String
Dim rsInsts As ADODB.Recordset
    
    On Error GoTo ErrLabel
    
    Call cboSubject.Clear
    mlSelSubjectId = 0
    
    sSQL = "SELECT PersonId from TrialSubject"
    sSQL = sSQL & " WHERE ClinicalTrialId = " & mlSelStudyId
    sSQL = sSQL & " AND TrialSite = '" & msSelSite & "'"
    sSQL = sSQL & " ORDER BY PersonId"
  
    Set rsInsts = New ADODB.Recordset
    rsInsts.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If Not rsInsts.EOF Then rsInsts.MoveFirst
    
    Do While Not rsInsts.EOF
        ' Add subjectID to combo
        cboSubject.AddItem CStr(rsInsts.Fields("PersonId"))
        rsInsts.MoveNext
    Loop
        
    rsInsts.Close
    Set rsInsts = Nothing
    
    If cboSubject.ListCount > 0 Then
        cboSubject.ListIndex = 0
        Call cboSubject_Click
    End If

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.LoadSubjectIds"
    
End Sub

'--------------------------------------------
Private Sub cboSubject_Click()
'--------------------------------------------
' Select a subject
'--------------------------------------------
   
    If cboSubject.ListIndex > -1 Then
        If cboSubject.List(cboSubject.ListIndex) <> mlSelSubjectId Then
            mlSelSubjectId = cboSubject.List(cboSubject.ListIndex)
        End If
    Else
        If mlSelSubjectId > 0 Then
            mlSelSubjectId = 0
        End If
    End If
    Call RememberSubject

End Sub

''--------------------------------------------
'Private Sub mnuTAzConsistency_Click()
''--------------------------------------------
'' Perform integrity check
''--------------------------------------------
'Dim sDataFileName As String
'
'    On Error Resume Next
'    CommonDialog1.CancelError = True
'    CommonDialog1.Flags = cdlOFNFileMustExist
'    CommonDialog1.Filter = "Text (*.txt)|*.txt"
'    CommonDialog1.FilterIndex = 1
'    CommonDialog1.ShowOpen
'    'If a valid file has been selected then process it
'    If Err.Number <> cdlCancel Then
'        sDataFileName = CommonDialog1.FileTitle
'        txtMsg.Text = Format(Now, "dd/mm/yyyy hh:mm:ss") & vbCrLf
'        txtMsg.Text = txtMsg.Text & DataIntegrityReport(mlSelStudyId, msSelSite, mlSelSubjectId, sDataFileName)
'    End If
'
'End Sub

'--------------------------------------------
Private Sub mnuTAzConsistency_Click()
'--------------------------------------------
' Perform integrity check
' This looks at every data value in the AREZZO Patient State
' and checks that it exists in the DIR table
' Offers the option of deleting extra values
' NCJ 20 June 07 - Added subject locking and cache handling
'--------------------------------------------
Dim sDataFileName As String
Dim sQuery As String
Dim sR As String
Dim sDelFileName As String
Dim bStuffToDelete As Boolean
Dim sReport As String
Dim sMSG As String
Dim sLockToken As String
Dim sCacheToken As String
Dim sLockMsg As String
Dim bReload As Boolean
Dim sVBErr As String

    On Error GoTo ErrLabel
    
    HourglassOn
    
    ' Get a cache token so we can check that it hasn't changed later
    sCacheToken = MACROLOCKBS30.CacheAddSubjectRow(goUser.CurrentDBConString, mlSelStudyId, msSelSite, mlSelSubjectId)
    bStuffToDelete = False
    sLockToken = ""
    
    ' Get all the AREZZO data
    sDataFileName = CreateDataListing
    
    If sDataFileName = "" Then
        ' Something went wrong with PLM file
        sReport = "Subject: " & msSubjectSpec & vbCrLf & gsPLM_ERROR
    Else
        sReport = "Subject: " & msSubjectSpec & vbCrLf _
                & DataIntegrityReport(mlSelStudyId, msSelSite, mlSelSubjectId, _
                                        sDataFileName, sDelFileName, bStuffToDelete)
    End If
    
    Call ShowTime
    Call ShowMessage(sReport)
    
    ' Offer to delete as long as it's not from a remote site
    If bStuffToDelete And Not mbSelSiteRemote Then
        If DialogQuestion("Do you want to delete these extra AREZZO data values?") = vbYes Then
            ' Make absolutely sure!
            sMSG = "This will delete all the listed values from this subject's AREZZO patient state." & vbCrLf _
                & "This is irreversible." & vbCrLf _
                & "Are you sure you want to do this?"
            If DialogQuestion(sMSG) = vbYes Then
                ' Do the deletes based on the "deletions" file
                ' See if we can lock the subject
                If LockForSave(sCacheToken, sLockMsg, sLockToken, bReload, mlSelStudyId, msSelSite, mlSelSubjectId) Then
                    If Not bReload Then
                        Call ShowMessage("Deleting data values from subject " & msSubjectSpec)
                        If DoTheDeletes(moArezzo, sDelFileName) Then
                            ' Save the subject
                            Call SaveAREZZOSubject(moArezzo, mlSelStudyId, msSelSite, mlSelSubjectId)
                            ' Invalidate the Cache to let others know it's changed
                            Call MACROLOCKBS30.CacheInvalidate(goUser.CurrentDBConString, mlSelStudyId, msSelSite, mlSelSubjectId, sCacheToken)
                            Call ShowMessage("Values deleted from subject " & msSubjectSpec)
                        Else
                            Call ShowMessage("Error: Unable to delete values from subject " & msSubjectSpec)
                        End If
                    Else
                        ' Someone else has got in there and changed things!
                        Call DialogError(" Another user has changed this subject. Please try again.")
                        Call ShowMessage("Another user has edited this subject - unable to delete values")
                    End If
                    Call MACROLOCKBS30.UnlockSubjectForSaving(goUser.CurrentDBConString, sLockToken, mlSelStudyId, msSelSite, mlSelSubjectId)
                Else
                    ' We didn't get the subject lock
                    sMSG = "Unable to save subject" & vbCrLf & sLockMsg
                    Call DialogError(sMSG)
                    Call ShowMessage("Unable to save - " & sLockMsg)

                End If
            End If
        End If
    End If
    
    ' We've finished now
    Call MACROLOCKBS30.CacheRemoveSubjectRow(goUser.CurrentDBConString, sCacheToken)
    
    HourglassOff

Exit Sub
ErrLabel:
    On Error Resume Next
    ' Save the error details
    sVBErr = Err.Number & " - " & Err.Description
    ' Unlock subject if necessary
    If sLockToken <> "" Then
        Call MACROLOCKBS30.UnlockSubjectForSaving(goUser.CurrentDBConString, sLockToken, mlSelStudyId, msSelSite, mlSelSubjectId)
    End If
    DialogError ("An error occurred." & vbCrLf & sVBErr)
    
End Sub

'--------------------------------------------------------------------
Private Function CreateDataListing() As String
'--------------------------------------------------------------------
' Load the currently selected subject and create an AREZZO data listing file
' Return file name of created file, or "" if error in PLM load
'--------------------------------------------------------------------
Dim sDataFileName As String
Dim sQuery As String
Dim sR As String

    On Error GoTo ErrLabel
    
    sDataFileName = ""
    
    ' Load the chosen subject into AREZZO
    If LoadAREZZOSubject(moArezzo, msSelStudyName, mlSelStudyId, msSelSite, mlSelSubjectId) Then
        ' Create an appropriate csv file name
        sDataFileName = gsTEMP_PATH & msSelStudyName & msSelSite & mlSelSubjectId & "_Data_" & _
                            Format(Now, "yyyymmddhhmmss") & ".csv"
        sQuery = "alldatarefs( '" & sDataFileName & "' ). "
        Call moArezzo.ALM.GetPrologResult(sQuery, sR)
    End If
    
    CreateDataListing = sDataFileName
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.CreateDataListing"

End Function

'--------------------------------------------------------------------
Private Function CreateTasksListing() As String
'--------------------------------------------------------------------
' Load the currently selected subject and create an AREZZO tasks listing file
' Return file name of created file, or "" if error in PLM load
'--------------------------------------------------------------------
Dim sDataFileName As String
Dim sQuery As String
Dim sR As String

    On Error GoTo ErrLabel
    
    sDataFileName = ""
    
    ' Load the chosen subject into AREZZO
    If LoadAREZZOSubject(moArezzo, msSelStudyName, mlSelStudyId, msSelSite, mlSelSubjectId) Then
        ' Create an appropriate csv file name
        sDataFileName = gsTEMP_PATH & msSelStudyName & msSelSite & mlSelSubjectId & "_Tasks_" & _
                            Format(Now, "yyyymmddhhmmss") & ".csv"
        sQuery = "alltasks( '" & sDataFileName & "' ). "
        Call moArezzo.ALM.GetPrologResult(sQuery, sR)
    End If
    
    CreateTasksListing = sDataFileName
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmMenu.CreateTasksListing"

End Function

'--------------------------------------------------------------------
Private Sub mnuTCLMMemory_Click()
'--------------------------------------------------------------------
' Report on memory situation when CLM file is loaded
'--------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    HourglassOn
    
    If msSelStudyName > "" Then
        ShowTime
        Call ShowMessage(CLMMemory(msSelStudyName, moArezzo))
    End If
    
    HourglassOff

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "mnuTCLMMemory_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Private Sub mnuTPLMMemory_Click()
'--------------------------------------------------------------------
' Report on memory situation when CLM and PLM files are loaded
'--------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    HourglassOn
    
    If msSubjectSpec > "" Then
        ShowTime
        Call ShowMessage(PLMMemory(msSubjectSpec, msSelStudyName, mlSelStudyId, msSelSite, mlSelSubjectId, moArezzo))
    End If
    
    HourglassOff

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "mnuTPLMMemory_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Private Sub mnuTCLMReport_Click()
'--------------------------------------------------------------------
' Run the CLM report (Count internal triggers etc.)
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    HourglassOn
    
    If msSelStudyName > "" Then
        ShowTime
        Call ShowMessage(CLMAnalyse(msSelStudyName, moArezzo))
    End If
    
    HourglassOff

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "mnuTCLMReport_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Private Sub mnuFDataReport_Click()
'--------------------------------------------------------------------
' Save a file containing data listing from AREZZO Patient State
'--------------------------------------------------------------------
Dim sFile As String
Dim sReport As String

    On Error GoTo ErrHandler
    
    sReport = "Subject: " & msSubjectSpec & vbCrLf
    sFile = CreateDataListing
    If sFile > "" Then
        sReport = sReport & _
                    "Subject data file saved as: " & sFile
    Else
        sReport = sReport & gsPLM_ERROR
    End If
    Call ShowTime
    Call ShowMessage(sReport)

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "mnuFDataReport_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Private Sub mnuTDiagnose_Click()
'--------------------------------------------------------------------
' Run the integrity report on current subject (look for orphaned references etc.)
'--------------------------------------------------------------------
Dim sReport As String

    On Error GoTo ErrHandler
    
    HourglassOn
    
    ShowTime
    sReport = "Subject: " & msSubjectSpec & vbCrLf & _
                PatIntegrityReport(mlSelStudyId, msSelSite, mlSelSubjectId)
    Call ShowMessage(sReport)
    
    HourglassOff

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "mnuTDiagnose_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Private Sub mnuTRecreate_Click()
'--------------------------------------------------------------------
' Recreate the AREZZO patient state for the currently selected subject
'--------------------------------------------------------------------
Dim sConfMsg As String
Dim oReconstruct As Reconstructor
Dim slogFileName As String

    On Error GoTo ErrHandler
    
    sConfMsg = "This will recreate the AREZZO patient state for subject " & vbCrLf _
            & msSubjectSpec & vbCrLf _
            & "Are you sure you wish to continue?"
    If DialogQuestion(sConfMsg) = vbYes Then
'        DialogInformation "Sorry, this functionality not implemented yet!"
        Call ShowTime
        Call ShowMessage("Reconstructing subject " & msSubjectSpec)
        Call HourglassOn
        
        slogFileName = gsTEMP_PATH & "Reconstruct_" & Format(Now, "yyyymmdd_hhmmss") & ".log"
        Set oReconstruct = New Reconstructor
        Call oReconstruct.InitReconstruction(slogFileName, goUser)
        Call oReconstruct.Reconstruct(mlSelStudyId, msSelSite, mlSelSubjectId, msSelStudyName, "", moArezzo)
        Call oReconstruct.EndReconstruction
        Set oReconstruct = Nothing
        
        Call HourglassOff
        Call ShowMessage("Reconstruction complete. Log file saved as: " & slogFileName)
    End If

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "mnuTRecreate_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Private Sub mnuFTasksReport_Click()
'--------------------------------------------------------------------
' Create a file containing listing of task instances from patient state
'--------------------------------------------------------------------
Dim sFile As String
Dim sReport As String

    On Error GoTo ErrHandler
    
    sReport = "Subject: " & msSubjectSpec & vbCrLf
    sFile = CreateTasksListing
    If sFile = "" Then
        sReport = sReport & gsPLM_ERROR
    Else
        sReport = sReport & _
                    "Subject tasks file saved as: " & sFile
    End If
    Call ShowTime
    Call ShowMessage(sReport)
    
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "mnuFTasksReport_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Private Function LockForSave(sCacheToken As String, ByRef sLockErrMsg As String, ByRef sToken As String, ByRef bNeedToReloadSubject As Boolean, _
                                lStudyId As Long, sSite As String, lSubjectId As Long) As Boolean
'---------------------------------------------------------------------
' Lock a subject for saving the patient state.
' NCJ 20 Jun 07 - Copied from DEBS
' Returns True/False for success/failure
'   sToken: token if lock successful or "" if not
'   sLockErrMsg: the reason the lock failed or "" if successful
'   nNeedToReloadSubject: if the lock is successful this returns whehther the subject needs reloading
'---------------------------------------------------------------------
Dim sLockDetails As String
Dim sCon As String
Dim sUser As String

Const sSTUDY_BEING = "This study is currently being "
Const sSUBJECT_BEING = "This subject is currently being "
Const sEDITED = "edited by "
Const sSAVED = "saved by "
Const sANOTHER_USER = "another user"

    On Error GoTo Errorlabel
    
    'set initial output variables to failure (change it later if we have success)
    sToken = ""
    sLockErrMsg = ""
    LockForSave = False
    
    ' The DB connection
    sCon = goUser.CurrentDBConString
    sUser = goUser.UserName
    
    sToken = MACROLOCKBS30.LockSubjectForSaving(sCon, sUser, lStudyId, sSite, lSubjectId)
    Select Case sToken
    Case MACROLOCKBS30.DBLocked.dblStudy
        sLockDetails = MACROLOCKBS30.LockDetailsStudy(sCon, lStudyId)
        sLockErrMsg = sSTUDY_BEING & sEDITED
        If sLockDetails = "" Then
            sLockErrMsg = sLockErrMsg & sANOTHER_USER
        Else
            sLockErrMsg = sLockErrMsg & Split(sLockDetails, "|")(0) & "."
        End If
        sToken = ""
    Case MACROLOCKBS30.DBLocked.dblSubject
        sLockDetails = MACROLOCKBS30.LockDetailsSubject(sCon, lStudyId, sSite, lSubjectId)
        sLockErrMsg = sSUBJECT_BEING & sEDITED
        If sLockDetails = "" Then
            sLockErrMsg = sLockErrMsg & sANOTHER_USER
        Else
            sLockErrMsg = sLockErrMsg & Split(sLockDetails, "|")(0) & "."
        End If
        sToken = ""
    Case MACROLOCKBS30.DBLocked.dblEFormInstance
        sLockDetails = MACROLOCKBS30.LockDetailsSubjectSave(sCon, lStudyId, sSite, lSubjectId)
        sLockErrMsg = sSUBJECT_BEING & sSAVED
        If sLockDetails = "" Then
            sLockErrMsg = sLockErrMsg & sANOTHER_USER
        Else
            sLockErrMsg = sLockErrMsg & Split(sLockDetails, "|")(0) & "."
        End If
        sToken = ""
    Case Else
        'we have a lock, but we need to check whether cache is invalid to decide whether to reload or not
        bNeedToReloadSubject = Not MACROLOCKBS30.CacheEntryStillValid(sCon, sCacheToken)
        'function to return success
        LockForSave = True
    End Select
    Exit Function
    
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmMenu.LockForSave"

End Function
