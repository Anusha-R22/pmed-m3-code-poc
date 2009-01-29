VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDataVerification 
   Caption         =   "MACRO Double Data Verification"
   ClientHeight    =   8775
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   9405
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picFontSizing 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      ScaleHeight     =   555
      ScaleWidth      =   1275
      TabIndex        =   17
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame fraVSD 
      Caption         =   "Verification Session Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4200
      Left            =   100
      TabIndex        =   7
      Top             =   200
      Width           =   7700
      Begin VB.CommandButton cmdExittVerify 
         Caption         =   "Exit Verification"
         Height          =   375
         Left            =   3480
         TabIndex        =   19
         Top             =   3720
         Width           =   1575
      End
      Begin VB.CommandButton cmdRestartVerify 
         Caption         =   "Restart Verification"
         Height          =   375
         Left            =   1800
         TabIndex        =   18
         Top             =   3720
         Width           =   1575
      End
      Begin VB.CommandButton cmdVerify 
         Caption         =   "Run Verification"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   3720
         Width           =   1575
      End
      Begin VB.ComboBox cboSite 
         Height          =   315
         Left            =   6100
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   450
         Width           =   1500
      End
      Begin VB.ComboBox cboStudy 
         Height          =   315
         Left            =   3900
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   450
         Width           =   2000
      End
      Begin MSComctlLib.ListView lvwSubjectVisits 
         Height          =   2460
         Left            =   100
         TabIndex        =   15
         Top             =   1155
         Width           =   7500
         _ExtentX        =   13229
         _ExtentY        =   4339
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblSite 
         Caption         =   "Site"
         Height          =   195
         Left            =   6100
         TabIndex        =   12
         Top             =   200
         Width           =   300
      End
      Begin VB.Label lblStudy 
         Caption         =   "Study"
         Height          =   255
         Left            =   3900
         TabIndex        =   10
         Top             =   200
         Width           =   405
      End
      Begin VB.Label Label1 
         Caption         =   "Select the Study and Site of the data to be verified."
         Height          =   255
         Left            =   100
         TabIndex        =   9
         Top             =   500
         Width           =   3720
      End
      Begin VB.Label Label2 
         Caption         =   "Select the Subject/Visit entries to be verified."
         Height          =   255
         Left            =   100
         TabIndex        =   8
         Top             =   900
         Width           =   3255
      End
   End
   Begin VB.PictureBox picVer 
      Height          =   4095
      Left            =   100
      ScaleHeight     =   4035
      ScaleWidth      =   7635
      TabIndex        =   0
      Top             =   4500
      Width           =   7700
      Begin VB.HScrollBar HScroll 
         Height          =   300
         LargeChange     =   900
         Left            =   120
         SmallChange     =   30
         TabIndex        =   6
         Top             =   3000
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.VScrollBar VScroll 
         Height          =   2385
         LargeChange     =   900
         Left            =   6960
         SmallChange     =   30
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox picVerEForm 
         BackColor       =   &H00FFFFFF&
         Height          =   2445
         Left            =   0
         ScaleHeight     =   2385
         ScaleWidth      =   5700
         TabIndex        =   1
         Top             =   0
         Width           =   5760
         Begin VB.TextBox txtResponsePass2 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   3000
            TabIndex        =   16
            Top             =   480
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtResponsePass1 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label lblUserPass2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "User Pass 2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   4200
            TabIndex        =   21
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblUserPass1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "User Pass 1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   1680
            TabIndex        =   20
            Top             =   480
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lblQuestionLabel 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.Label lblHeader 
            BackColor       =   &H8000000E&
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   0
            Width           =   2535
         End
      End
   End
End
Attribute VB_Name = "frmDataVerification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
' File:         frmDataVerification.frm
' Copyright:    InferMed Ltd. 2007. All Rights Reserved
' Author:       Mo Morris, September 2007
' Purpose:      Contains form that is used to run MACRO Double Data Entry Verification Process
'----------------------------------------------------------------------------------------'
'   Revisions:
'----------------------------------------------------------------------------------------'

Option Explicit

Private mlVerTrialId As Long
Private msVerSite As String
Private msVerTrialName As String
Private mbVerifyHalted As Boolean
Private mcolForVerification As Collection

'--------------------------------------------------------------------
Private Sub LoadStudyCombo()
'--------------------------------------------------------------------
Dim colstudies As Collection
Dim oStudy As Study

    On Error GoTo ErrHandler

    'Clear current contents of cboStudy
    cboStudy.Clear
    glSelTrialId = 0
    
    Set colstudies = goUser.GetNewSubjectStudies

    For Each oStudy In colstudies
        cboStudy.AddItem oStudy.StudyName
        cboStudy.ItemData(cboStudy.NewIndex) = oStudy.StudyId
    Next

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "LoadStudyCombo")
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
Private Sub LoadSiteCombo()
'--------------------------------------------------------------------
Dim colSites As Collection
Dim oSite As Site

    On Error GoTo ErrHandler
    
    'Clear current contents of cboSite
    cboSite.Clear
    gsSelSite = ""
    
    Set colSites = goUser.GetNewSubjectSites(mlVerTrialId)
    
    For Each oSite In colSites
        cboSite.AddItem oSite.Site
    Next
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "LoadSiteCombo")
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
Private Sub cboSite_Click()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    'exit if no item is currently selected
    If cboSite.ListIndex = -1 Then Exit Sub
    
    msVerSite = Trim(cboSite.Text)
    
    Call GetSubjectsToVerify

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cboSite_Click")
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
Private Sub cboStudy_Click()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler

    'exit if no item is currently selected
    If cboStudy.ListIndex = -1 Then Exit Sub
    
    msVerTrialName = Trim(cboStudy.Text)
    mlVerTrialId = cboStudy.ItemData(cboStudy.ListIndex)

    Call LoadSiteCombo

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cboStudy_Click")
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
Private Sub cmdExittVerify_Click()
'--------------------------------------------------------------------

    Call LaunchBatchDataUpload

    Unload Me

End Sub

'--------------------------------------------------------------------
Private Sub cmdRestartVerify_Click()
'--------------------------------------------------------------------
Dim i As Integer
Dim j As Integer
Dim lVerPersonId As Long
Dim lVerVisitId As Long
Dim nVerVisitCycleNumber As Integer
Dim lFirsteForm As Long
Dim lVisitDateeForm As Long
Dim lNexteForm As Long

    'Disable cmdRestartVerify
    cmdRestartVerify.Enabled = False

    mbVerifyHalted = False

    For i = 1 To mcolForVerification.Count
        lVerPersonId = lvwSubjectVisits.ListItems(mcolForVerification.Item(i)).SubItems(1)
        lVerVisitId = lvwSubjectVisits.ListItems(mcolForVerification.Item(i)).SubItems(4)
        nVerVisitCycleNumber = lvwSubjectVisits.ListItems(mcolForVerification.Item(i)).SubItems(3)
        lVisitDateeForm = GetVisitDateeForm(mlVerTrialId, lVerVisitId)
        lFirsteForm = GetFirsteFormInVisit(mlVerTrialId, lVerVisitId)
        'Loop through eForms that exist within the selected Study/Site/Subject/Visit.
        'Note that if the Visit contains an enterable Visit-Date-Question then
        'that Visit-Date-eForm will be processed first
        If lVisitDateeForm > 0 Then
            Call VerifyeForms(lVerPersonId, lVerVisitId, nVerVisitCycleNumber, lVisitDateeForm)
        End If
        If Not mbVerifyHalted Then
            Call VerifyeForms(lVerPersonId, lVerVisitId, nVerVisitCycleNumber, lFirsteForm)
            If Not mbVerifyHalted Then
                Do Until Not NexteFormExists(mlVerTrialId, lVerVisitId, lFirsteForm, lNexteForm)
                    lFirsteForm = lNexteForm
                    Call VerifyeForms(lVerPersonId, lVerVisitId, nVerVisitCycleNumber, lFirsteForm)
                    If mbVerifyHalted Then Exit Do
                Loop
            End If
        End If
        If mbVerifyHalted Then Exit For
    Next
    
    If Not mbVerifyHalted Then
        Call ClearVereForm
        picVerEForm.Height = 0
        picVerEForm.Width = 0
        picVerEForm.Visible = False
        'enable lvwSubjectVisits
        lvwSubjectVisits.Enabled = True
        'Call GetSubjectsToVerify for the purpose of appearing to clear processed entries
        Call GetSubjectsToVerify
    End If
    
    Call LaunchBatchDataUpload

End Sub

'--------------------------------------------------------------------
Private Sub cmdVerify_Click()
'--------------------------------------------------------------------
Dim i As Integer
Dim j As Integer
Dim lVerPersonId As Long
Dim lVerVisitId As Long
Dim nVerVisitCycleNumber As Integer
Dim lFirsteForm As Long
Dim lVisitDateeForm As Long
Dim lNexteForm As Long

    'disable cmdVerify
    cmdVerify.Enabled = False
    'disable lvwSubjectVisits
    lvwSubjectVisits.Enabled = False
    'Clear collection of selected lvwSubjectVisits entries to be verifed
    Set mcolForVerification = Nothing
    Set mcolForVerification = New Collection
    
    'Place the selected lvwSubjectVisits entries into mcolForVerification
    'becasue during the verification process lvwSubjectVisits entries become unselected
    For i = 1 To lvwSubjectVisits.ListItems.Count
        If lvwSubjectVisits.ListItems(i).Selected Then
            mcolForVerification.Add i, Str(i)
        End If
    Next

    mbVerifyHalted = False
    
    For i = 1 To mcolForVerification.Count
        lVerPersonId = lvwSubjectVisits.ListItems(mcolForVerification.Item(i)).SubItems(1)
        lVerVisitId = lvwSubjectVisits.ListItems(mcolForVerification.Item(i)).SubItems(4)
        nVerVisitCycleNumber = lvwSubjectVisits.ListItems(mcolForVerification.Item(i)).SubItems(3)
        lVisitDateeForm = GetVisitDateeForm(mlVerTrialId, lVerVisitId)
        lFirsteForm = GetFirsteFormInVisit(mlVerTrialId, lVerVisitId)
        'Loop through eForms that exist within the selected Study/Site/Subject/Visit.
        'Note that if the Visit contains an enterable Visit-Date-Question then
        'that Visit-Date-eForm will be processed first
        If lVisitDateeForm > 0 Then
            Call VerifyeForms(lVerPersonId, lVerVisitId, nVerVisitCycleNumber, lVisitDateeForm)
        End If
        If Not mbVerifyHalted Then
            Call VerifyeForms(lVerPersonId, lVerVisitId, nVerVisitCycleNumber, lFirsteForm)
            If Not mbVerifyHalted Then
                Do Until Not NexteFormExists(mlVerTrialId, lVerVisitId, lFirsteForm, lNexteForm)
                    lFirsteForm = lNexteForm
                    Call VerifyeForms(lVerPersonId, lVerVisitId, nVerVisitCycleNumber, lFirsteForm)
                    If mbVerifyHalted Then Exit Do
                Loop
            End If
        End If
        If mbVerifyHalted Then Exit For
    Next
    
    If Not mbVerifyHalted Then
        Call ClearVereForm
        picVerEForm.Height = 0
        picVerEForm.Width = 0
        picVerEForm.Visible = False
        'enable lvwSubjectVisits
        lvwSubjectVisits.Enabled = True
        'Call GetSubjectsToVerify for the purpose of appearing to clear processed entries
        Call GetSubjectsToVerify
    End If
    
    Call LaunchBatchDataUpload

End Sub

'--------------------------------------------------------------------
Private Sub Form_Load()
'--------------------------------------------------------------------
Dim colmX As MSComctlLib.ColumnHeader

    Call LoadStudyCombo
    
    Set colmX = lvwSubjectVisits.ColumnHeaders.Add(, , "Subject Label", 1500)
    Set colmX = lvwSubjectVisits.ColumnHeaders.Add(, , "PersonId", 1000)
    Set colmX = lvwSubjectVisits.ColumnHeaders.Add(, , "Visit Name (Code)", 2000)
    Set colmX = lvwSubjectVisits.ColumnHeaders.Add(, , "Visit Cycle", 1000)
    Set colmX = lvwSubjectVisits.ColumnHeaders.Add(, , "VisitId", 0)
    
    cmdVerify.Enabled = False
    cmdRestartVerify.Enabled = False

    picVerEForm.Visible = False

End Sub

'--------------------------------------------------------------------
Private Function GetSubjectsToVerify() As ADODB.Recordset
'--------------------------------------------------------------------
Dim sSQL As String
Dim sSQL2 As String
Dim rsPersonVisits As ADODB.Recordset
Dim rsPassCount As ADODB.Recordset
Dim itmX As MSComctlLib.ListItem
Dim i As Integer

    On Error GoTo ErrHandler
    
    'clear current content of SubjectVisits to verify
    lvwSubjectVisits.ListItems.Clear
    'disable the Run Verification command button
    cmdVerify.Enabled = False

    sSQL = "SELECT DISTINCT PersonId, VisitId, VisitCycleNumber FROM DoubleData " _
        & " WHERE ClinicalTrialId = " & mlVerTrialId _
        & " AND TrialSite = '" & msVerSite & "'" _
        & " AND Status = " & eDoubleDataStatus.Entered

    Set rsPersonVisits = New ADODB.Recordset
    rsPersonVisits.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsPersonVisits.RecordCount > 0 Then
        Do Until rsPersonVisits.EOF
            'Query for First Pass entries
            sSQL2 = "SELECT Count(*) as Pass1Count FROM DoubleData " _
                & " WHERE ClinicalTrialId = " & mlVerTrialId _
                & " AND TrialSite = '" & msVerSite & "'" _
                & " AND PersonId = " & rsPersonVisits!PersonID _
                & " AND VisitId = " & rsPersonVisits!VisitId _
                & " AND VisitCycleNumber = " & rsPersonVisits!VisitCycleNumber _
                & " AND Status = " & eDoubleDataStatus.Entered _
                & " AND Pass = 1"
            Set rsPassCount = New ADODB.Recordset
            rsPassCount.Open sSQL2, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText

            If rsPassCount!Pass1Count > 0 Then
                'Query for Second Pass entries
                sSQL2 = "SELECT Count(*) as Pass2Count FROM DoubleData " _
                    & " WHERE ClinicalTrialId = " & mlVerTrialId _
                    & " AND TrialSite = '" & msVerSite & "'" _
                    & " AND PersonId = " & rsPersonVisits!PersonID _
                    & " AND VisitId = " & rsPersonVisits!VisitId _
                    & " AND VisitCycleNumber = " & rsPersonVisits!VisitCycleNumber _
                    & " AND Status = " & eDoubleDataStatus.Entered _
                    & " AND Pass = 2"
                Set rsPassCount = New ADODB.Recordset
                rsPassCount.Open sSQL2, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
                
                If rsPassCount!Pass2Count > 0 Then
                    Set itmX = lvwSubjectVisits.ListItems.Add(, , SubjectLabelFromTrialSiteId(mlVerTrialId, msVerSite, rsPersonVisits!PersonID))
                    itmX.SubItems(1) = rsPersonVisits!PersonID
                    itmX.SubItems(2) = VisitNameFromId(mlVerTrialId, rsPersonVisits!VisitId, 1) & " (" & VisitCodeFromId(mlVerTrialId, rsPersonVisits!VisitId) & ")"
                    itmX.SubItems(3) = rsPersonVisits!VisitCycleNumber
                    itmX.SubItems(4) = rsPersonVisits!VisitId
                End If
            End If
            
            rsPersonVisits.MoveNext
        Loop
    
        rsPassCount.Close
        Set rsPassCount = Nothing
        
    End If
    
    rsPersonVisits.Close
    Set rsPersonVisits = Nothing
    
    For i = 1 To 4
        Call lvw_SetColWidth(lvwSubjectVisits, i, LVSCW_AUTOSIZE_USEHEADER)
    Next i

Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "GetSubjectsToVerify")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Function

'---------------------------------------------------------------------
Private Function ListViewSelectedCount() As Long
'---------------------------------------------------------------------
' NCJ 9 Mar 04 - Roche UAT Bug 44 - Change nCount as integer to lCount as long
'---------------------------------------------------------------------
Dim olvwItem As ListItem
Dim lCount As Long

    On Error GoTo ErrHandler
    
    lCount = 0
    For Each olvwItem In lvwSubjectVisits.ListItems
        If olvwItem.Selected Then lCount = lCount + 1
    Next
    
    ListViewSelectedCount = lCount

Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "ListViewSelectedCount")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Function

'--------------------------------------------------------------------
Private Sub Form_Resize()
'--------------------------------------------------------------------
Dim nFormWidth As Integer
Dim nFormHeight As Integer

    On Error GoTo ErrHandler

    'check that form has not been minimised
    If Me.WindowState = vbMinimized Then Exit Sub
    
    'force a minimum hieght for the form
    If Me.Height < mnMINVFORMHEIGHT Then Me.Height = mnMINVFORMHEIGHT
    
    'force a minimum width for the form
    If Me.Width < mnMINVFORMWIDTH Then Me.Width = mnMINVFORMWIDTH
    
    nFormWidth = Me.Width
    nFormHeight = Me.Height

    picVer.Width = nFormWidth - 300
    picVer.Height = nFormHeight - 5000
    picVerEForm.Left = 0
    
    'Assess the need for vertical scrollbar
    If picVerEForm.Height > picVer.Height Then
        Set VScroll.Container = picVer
        VScroll.Max = (picVerEForm.Height - picVer.Height) / 10
        VScroll.Top = 0
        If picVerEForm.Width > picVer.Width Then
            VScroll.Height = picVer.ScaleHeight - HScroll.Height
        Else
            VScroll.Height = picVer.ScaleHeight
        End If
        VScroll.Left = picVer.ScaleWidth - VScroll.Width
        VScroll.LargeChange = CInt(((VScroll.Max / 2) / 10) + 1)
        VScroll.SmallChange = CInt(((VScroll.Max / 10) / 10) + 1)
        VScroll.Value = 0
        VScroll.Visible = True
    Else
        VScroll.Visible = False
    End If
    
    'Assess the need for horizontal scrollbar
    If picVerEForm.Width > picVer.Width Then
        Set HScroll.Container = picVer
        HScroll.Max = (picVerEForm.Width - picVer.Width) / 10
        HScroll.Top = picVer.ScaleHeight - HScroll.Height
        If VScroll.Visible = True Then
            HScroll.Width = picVer.ScaleWidth - VScroll.Width
        Else
            HScroll.Width = picVer.ScaleWidth
        End If
        HScroll.Left = 0
        HScroll.LargeChange = CInt(((HScroll.Max / 2) / 10) + 1)
        HScroll.SmallChange = CInt(((HScroll.Max / 10) / 10) + 1)
        HScroll.Value = 0
        HScroll.Visible = True
    Else
        HScroll.Visible = False
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

    Unload Me

End Sub

'--------------------------------------------------------------------
Private Sub HScroll_Change()
'--------------------------------------------------------------------

    picVerEForm.Left = CSng(-HScroll.Value) * 10

End Sub

'--------------------------------------------------------------------
Private Sub HScroll_Scroll()
'--------------------------------------------------------------------

    picVerEForm.Left = CSng(-HScroll.Value) * 10

End Sub

'--------------------------------------------------------------------
Private Sub lvwSubjectVisits_ItemClick(ByVal Item As MSComctlLib.ListItem)
'--------------------------------------------------------------------

    On Error GoTo ErrHandler

    If ListViewSelectedCount = 0 Then
        cmdVerify.Enabled = False
    Else
        cmdVerify.Enabled = True
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "lvwBuffer_ItemClick")
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
Private Function GetSubjectData(ByVal lPersonId As Long, _
                                ByVal lVisitId As Long) As ADODB.Recordset
'--------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    sSQL = "SELECT * " _
        & " FROM DoubleData, CRFElement " _
        & " WHERE DoubleData.ClinicalTrialId = CRFElement.ClinicalTrialId " _
        & " AND DoubleData.CRFPageId = CRFElement.CRFPageId " _
        & " AND DoubleData.DataItemId = CRFElement.DataItemId " _
        & " AND DoubleData.ClinicalTrialId = " & mlVerTrialId _
        & " AND DoubleData.TrialSite = '" & msVerSite & "'" _
        & " AND DoubleData.PersonId = " & lPersonId _
        & " AND DoubleData.Visitid = " & lVisitId _
        & " AND DoubleData.Status = " & eDoubleDataStatus.Entered _
        & " ORDER BY DoubleData.CRFPageId, CRFElement.FieldOrder, DoubleData.RepeatNumber, CRFElement.QGroupFieldOrder"
    
    Set GetSubjectData = New ADODB.Recordset
    GetSubjectData.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText

Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "GetSubjectData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Function

'--------------------------------------------------------------------
Private Function GetFirsteFormInVisit(ByVal lClinicalTrialId, _
                                        ByVal lVisitId As Long) As Long
'--------------------------------------------------------------------
Dim sSQL As String
Dim rsEForm As ADODB.Recordset
Dim lCRFPageId As Long

    sSQL = "SELECT CRFPage.CRFPageId " _
        & "FROM CRFPage, StudyVisitCRFPage " _
        & "WHERE CRFPage.ClinicalTrialId = StudyVisitCRFPage.ClinicalTrialId " _
        & "AND CRFPage.CRFPageId = StudyVisitCRFPage.CRFPageId " _
        & "AND CRFPage.ClinicalTrialId = " & lClinicalTrialId & " " _
        & "AND StudyVisitCRFPage.VisitId = " & lVisitId & " " _
        & "AND StudyVisitCRFPage.eFormUse = 0 " _
        & "ORDER BY CRFPage.CRFPageOrder"
        
    Set rsEForm = New ADODB.Recordset
    rsEForm.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If rsEForm.RecordCount = 0 Then
        'This should not happen, a visit in an active study should always include at least one eForm
        lCRFPageId = 0
    Else
        lCRFPageId = rsEForm!CRFPageId
    End If
    
    rsEForm.Close
    Set rsEForm = Nothing
    
    GetFirsteFormInVisit = lCRFPageId

End Function

'--------------------------------------------------------------------
Private Function GetVisitDateeForm(ByVal lClinicalTrialId, _
                                    ByVal lVisitId As Long) As Long
'--------------------------------------------------------------------
'This function checks to see if a Visit contains a Visit-Date-eForm.
'If no Visit-Date-eForm exists the function returns 0.
'If a Visit-Date-eForm exists then the DataItemId of its Visit-Date-Question is retrieved.
'The resulting DataItemId is then checked for being derived.
'If the Visit-Date-Question is derived the function returns 0.
'If a Visit-Date-eForm with an enterable Visit-Date-Question exists the function returns the CRFPageId of the Visit-Date-eForm.
'--------------------------------------------------------------------
Dim sSQL As String
Dim rsEForm As ADODB.Recordset
Dim lVisitDateeFormId As Long
Dim lVisitDateQuestionId As Long

    'Query for a Visit-Date-eForm existing in specified Visit
    sSQL = "SELECT CRFPageId " _
        & "FROM StudyVisitCRFPage " _
        & "WHERE ClinicalTrialId = " & lClinicalTrialId & " " _
        & "AND VisitId = " & lVisitId & " " _
        & "AND eFormUse = 1"
        
    Set rsEForm = New ADODB.Recordset
    rsEForm.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If rsEForm.RecordCount = 0 Then
        lVisitDateeFormId = 0
    Else
        lVisitDateeFormId = rsEForm!CRFPageId
    End If

    rsEForm.Close
    Set rsEForm = Nothing
    
    If lVisitDateeFormId > 0 Then
        'Query for the Visit-Date-eForm's Visit-Date-Question
        sSQL = "SELECT DataItemId " _
            & "FROM CRFElement " _
            & "WHERE ClinicalTrialId = " & lClinicalTrialId & " " _
            & "AND CRFPageId = " & lVisitDateeFormId & " " _
            & "AND ElementUse = 1"
            
        Set rsEForm = New ADODB.Recordset
        rsEForm.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        If rsEForm.RecordCount = 0 Then
            'This should not occur, a Visit-Date-eForm should always have a Visit-Date-Question
            lVisitDateQuestionId = 0
        Else
            lVisitDateQuestionId = rsEForm!DataItemId
        End If
        
        rsEForm.Close
        Set rsEForm = Nothing
        
        If lVisitDateQuestionId > 0 Then
            'Is the Visit-Date-Question derived or enterable
            If QuestionIsDerived(lClinicalTrialId, lVisitDateQuestionId) Then
                GetVisitDateeForm = 0
            Else
                GetVisitDateeForm = lVisitDateeFormId
            End If
        End If
    Else
        GetVisitDateeForm = 0
    End If
            
End Function

'--------------------------------------------------------------------
Private Sub VerifyeForms(ByVal lPersonId As Long, _
                        ByVal lVisitId As Long, _
                        ByVal nVisitCycleNumber As Integer, _
                        ByVal lCRFPageId As Long)
'--------------------------------------------------------------------
Dim sSQL As String
Dim rsCRFPageCycleNumbers As ADODB.Recordset
Dim i As Integer

    'Query for CRFPageCycleNumbers from the specified Trial/Person/Visit/VisitCycle/CRFPage
    'A no-repeating eForm will only cycle once
    'A repeating eForm might have more than one cycle
    sSQL = "SELECT DISTINCT CRFPageCycleNumber FROM DoubleData " _
        & " WHERE ClinicalTrialId = " & mlVerTrialId _
        & " AND TrialSite = '" & msVerSite & "'" _
        & " AND PersonId = " & lPersonId _
        & " AND VisitId = " & lVisitId _
        & " AND VisitCycleNumber = " & nVisitCycleNumber _
        & " AND CRFPageId = " & lCRFPageId _
        & " AND Status = " & eDoubleDataStatus.Entered
    
    Set rsCRFPageCycleNumbers = New ADODB.Recordset
    rsCRFPageCycleNumbers.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    'Loop on CRFPageCycleNumber and call VerifySingleeForm for each instance
    For i = 1 To rsCRFPageCycleNumbers.RecordCount
        Call VerifySingleeForm(lPersonId, lVisitId, nVisitCycleNumber, lCRFPageId, rsCRFPageCycleNumbers!CRFPageCycleNumber)
        If mbVerifyHalted Then Exit For
        rsCRFPageCycleNumbers.MoveNext
    Next
    
End Sub

'--------------------------------------------------------------------
Private Sub VerifySingleeForm(ByVal lPersonId As Long, _
                        ByVal lVisitId As Long, _
                        ByVal nVisitCycleNumber As Integer, _
                        ByVal lCRFPageId As Long, _
                        ByVal nCRFPageCycleNumber As Integer)
'--------------------------------------------------------------------
Dim sHeader As String
Dim lTopCount As Long
Dim sSQL As String
Dim rsEFormQuestions As ADODB.Recordset
Dim nTextBoxheight As Integer
Dim n50CharWidth As Integer
Dim nLabelIndex As Integer
Dim nTextBox1Index As Integer
Dim nTextBox2Index As Integer
Dim nUserLabel1Index As Integer
Dim nUserLabel2Index As Integer
Dim lMaxWidth As Long
Dim oLabel As Label
Dim oTextBox1 As TextBox
Dim oTextBox2 As TextBox
Dim oUserLabel1 As Label
Dim oUserLabel2 As Label
Dim sDataItemName As String
Dim sDataItemPrefix As String
Dim nDataItemType As Integer
Dim nDataItemLength As Integer
Dim sResponse1 As String
Dim sResponse2 As String
Dim bVerOk As Boolean
Dim sUserName1 As String
Dim sUserName2 As String

    Call ClearVereForm
    Call SetFontAndColour
    
    picVerEForm.Top = 0
    picVerEForm.Left = 0

    nLabelIndex = 1
    nTextBox1Index = 1
    nTextBox2Index = 1
    nUserLabel1Index = 1
    nUserLabel2Index = 1
    lTopCount = 100
    picFontSizing.FontSize = gnRegDDFontSize
    nTextBoxheight = picFontSizing.TextHeight("X")
    n50CharWidth = picFontSizing.TextWidth(String(50, "X"))
    
    'Create a Header Label
    sHeader = "Subject: " & msVerTrialName & "/" & msVerSite & "/" & lPersonId _
        & "    Visit: " & VisitCodeFromId(mlVerTrialId, lVisitId) & "[" & nVisitCycleNumber _
        & "]    eForm: " & CRFPageCodeFromId(mlVerTrialId, lCRFPageId) & "[" & nCRFPageCycleNumber & "]"
    
    'Write a page header to the Verification Form
    With lblHeader
        .Caption = sHeader
        .Top = lTopCount
        .Left = 100
        .Width = picFontSizing.TextWidth(sHeader & "                   ")
        .Height = (nTextBoxheight + 100)
        .Visible = True
        'give lMaxWidth an initial value based on width of lblHeader
        lMaxWidth = .Left + .Width
    End With
    
    'increment lTopCount
    lTopCount = lTopCount + (2 * nTextBoxheight)
    
    'Query for all questions on the single eForm specified by Trial/Person/Visit/VisitCycle/CRFPage/CRFPageCycleNumber
    sSQL = "SELECT DISTINCT DataItemId, FieldOrder, RepeatNumber, QGroupFieldOrder FROM DoubleData " _
        & " WHERE ClinicalTrialId = " & mlVerTrialId _
        & " AND TrialSite = '" & msVerSite & "'" _
        & " AND PersonId = " & lPersonId _
        & " AND VisitId = " & lVisitId _
        & " AND VisitCycleNumber = " & nVisitCycleNumber _
        & " AND CRFPageId = " & lCRFPageId _
        & " AND CRFPageCycleNumber = " & nCRFPageCycleNumber _
        & " AND Status = " & eDoubleDataStatus.Entered _
        & " ORDER BY FieldOrder, RepeatNumber, QGroupFieldOrder"
        
    Set rsEFormQuestions = New ADODB.Recordset
    rsEFormQuestions.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    bVerOk = True
    
    'Loop through the eforms questions, writing caption, pass 1 response, pass 2 response to the Verification Form and compare
    Do Until rsEFormQuestions.EOF
        Call GetDataItemDetails(mlVerTrialId, rsEFormQuestions!DataItemId, sDataItemName, nDataItemType, nDataItemLength)
        'Create Question Label
        Load Me.lblQuestionLabel(nLabelIndex)
        Set oLabel = Me.lblQuestionLabel(nLabelIndex)
        nLabelIndex = nLabelIndex + 1
        'Check for a question that is part of a repeating question group
        If rsEFormQuestions!QGroupFieldOrder > 0 Then
            sDataItemName = " [" & rsEFormQuestions!FieldOrder & "." & rsEFormQuestions!QGroupFieldOrder & "." & rsEFormQuestions!RepeatNumber & "] " & sDataItemName
        Else
            sDataItemName = " [" & rsEFormQuestions!FieldOrder & "] " & sDataItemName
        End If
        With oLabel
            .Caption = sDataItemName
            .Left = 100
            .Top = lTopCount
            .Width = picFontSizing.TextWidth(sDataItemName & "   ")
            .Height = (nTextBoxheight + 100)
            .Visible = True
        End With
        'Create Pass1 TextBox
        sResponse1 = GetResponse(lPersonId, lVisitId, nVisitCycleNumber, lCRFPageId, nCRFPageCycleNumber, rsEFormQuestions!DataItemId, rsEFormQuestions!RepeatNumber, 1, sUserName1)
        Load Me.txtResponsePass1(nTextBox1Index)
        Set oTextBox1 = Me.txtResponsePass1(nTextBox1Index)
        nTextBox1Index = nTextBox1Index + 1
        With oTextBox1
            .Text = sResponse1
            .Top = lTopCount
            .Left = n50CharWidth + 600
            .Width = picFontSizing.TextWidth(String(nDataItemLength + 3, "_"))
            .Height = nTextBoxheight + 100
            .Tag = lPersonId & "|" & lVisitId & "|" & nVisitCycleNumber & "|" & lCRFPageId & "|" & nCRFPageCycleNumber _
                & "|" & rsEFormQuestions!DataItemId & "|" & rsEFormQuestions!FieldOrder _
                & "|" & rsEFormQuestions!RepeatNumber & "|" & rsEFormQuestions!QGroupFieldOrder & "|" & nDataItemType
            .Visible = True
        End With
        'Create Pass1 User label
        Load Me.lblUserPass1(nUserLabel1Index)
        Set oUserLabel1 = Me.lblUserPass1(nUserLabel1Index)
        nUserLabel1Index = nUserLabel1Index + 1
        sUserName1 = "[" & sUserName1 & "]"
        With oUserLabel1
            .Caption = sUserName1
            .Left = n50CharWidth + 600 + oTextBox1.Width + 50
            .Top = lTopCount
            .Width = picFontSizing.TextWidth(sUserName1)
            .Height = nTextBoxheight + 100
            .Visible = True
        End With
        'Create Pass2 TextBox
        sResponse2 = GetResponse(lPersonId, lVisitId, nVisitCycleNumber, lCRFPageId, nCRFPageCycleNumber, rsEFormQuestions!DataItemId, rsEFormQuestions!RepeatNumber, 2, sUserName2)
        Load Me.txtResponsePass2(nTextBox2Index)
        Set oTextBox2 = Me.txtResponsePass2(nTextBox2Index)
        nTextBox2Index = nTextBox2Index + 1
        With oTextBox2
            .Text = sResponse2
            .Top = lTopCount
            .Left = n50CharWidth + 600 + oTextBox1.Width + 50 + oUserLabel1.Width + 200
            .Width = picFontSizing.TextWidth(String(nDataItemLength + 3, "_"))
            .Height = nTextBoxheight + 100
            .Tag = lPersonId & "|" & lVisitId & "|" & nVisitCycleNumber & "|" & lCRFPageId & "|" & nCRFPageCycleNumber _
                & "|" & rsEFormQuestions!DataItemId & "|" & rsEFormQuestions!FieldOrder _
                & "|" & rsEFormQuestions!RepeatNumber & "|" & rsEFormQuestions!QGroupFieldOrder & "|" & nDataItemType
            .Visible = True
        End With
        'Create Pass2 User label
        Load Me.lblUserPass2(nUserLabel2Index)
        Set oUserLabel2 = Me.lblUserPass2(nUserLabel2Index)
        nUserLabel2Index = nUserLabel2Index + 1
        sUserName2 = "[" & sUserName2 & "]"
        With oUserLabel2
            .Caption = sUserName2
            .Left = n50CharWidth + 600 + oTextBox1.Width + 50 + oUserLabel1.Width + 200 + oTextBox2.Width + 50
            .Top = lTopCount
            .Width = picFontSizing.TextWidth(sUserName2)
            .Height = nTextBoxheight + 100
            .Visible = True
            If lMaxWidth < (.Left + .Width) Then
                lMaxWidth = .Left + .Width
            End If
        End With
        
        'Increment lTopCount
        lTopCount = lTopCount + (2 * nTextBoxheight)
        
        picVerEForm.Height = lTopCount + 500
        picVerEForm.Width = lMaxWidth + 500
        picVerEForm.Visible = True
          
        'Compare Pass 1 and Pass 2 responses
        If (sResponse1 = "") And (sResponse2 = "") Then
            'Pass1 and Pass2 responses both blank. No BDE required. Set as Verified
            Call ChangeDDStatusToVerified(lPersonId, lVisitId, nVisitCycleNumber, lCRFPageId, nCRFPageCycleNumber, rsEFormQuestions!DataItemId, rsEFormQuestions!RepeatNumber)
        ElseIf sResponse1 = sResponse2 Then
            'Pass1 and Pass2 responses exist and are identical. Create BDE. Set as Verfified
            Call GenerateBatchDataEntry(lPersonId, lVisitId, nVisitCycleNumber, lCRFPageId, nCRFPageCycleNumber, rsEFormQuestions!DataItemId, rsEFormQuestions!RepeatNumber, sResponse1)
            Call ChangeDDStatusToVerified(lPersonId, lVisitId, nVisitCycleNumber, lCRFPageId, nCRFPageCycleNumber, rsEFormQuestions!DataItemId, rsEFormQuestions!RepeatNumber)
        Else
            'Either Pass1 or Pass2 response does not exist or they differ. No BDE required. Not set as Verfified
            'Set Verify Not Ok flag
            bVerOk = False
        End If
        
        If Not bVerOk Then
            'Create global Verification Halted flag
            mbVerifyHalted = True
            'enable cmdRestartVerify
            cmdRestartVerify.Enabled = True
            'enable the Pass1 and Pass2 verification textboxes
            oTextBox1.Enabled = True
            oTextBox2.Enabled = True
            'Set focus to txtResponsePass1
            oTextBox1.SetFocus
            Exit Do
        End If
        
        rsEFormQuestions.MoveNext
    Loop
    
    Call VereFormChecks
    Call VereFormScrollbars
 
End Sub

'--------------------------------------------------------------------
Private Sub ClearVereForm()
'--------------------------------------------------------------------
Dim oControl As Control

    picVerEForm.Visible = False
    VScroll.Visible = False
    HScroll.Visible = False

    'clear previous picVerEForm content
    For Each oControl In Me.Controls
        If oControl.Name = "lblQuestionLabel" Or oControl.Name = "txtResponsePass1" Or oControl.Name = "txtResponsePass2" _
            Or oControl.Name = "lblUserPass1" Or oControl.Name = "lblUserPass2" Then
            If oControl.Index > 0 Then
                Unload oControl
                DoEvents
            End If
        End If
    Next

End Sub

'--------------------------------------------------------------------
Private Sub SetFontAndColour()
'--------------------------------------------------------------------
Dim oControl As Control

    For Each oControl In Me.Controls
        If oControl.Name = "lblQuestionLabel" Or oControl.Name = "txtResponsePass1" Or oControl.Name = "txtResponsePass2" _
            Or oControl.Name = "lblHeader" Or oControl.Name = "lblUserPass1" Or oControl.Name = "lblUserPass2" Then
            oControl.FontSize = gnRegDDFontSize
            Select Case oControl.Name
            Case "txtResponsePass1", "txtResponsePass2"
                oControl.BackColor = glRegDDLightColour
            Case "lblHeader", "lblQuestionLabel", "lblUserPass1", "lblUserPass2"
                oControl.BackColor = glRegDDMediumColour
            End Select
        End If
    Next
    
    picVerEForm.BackColor = glRegDDMediumColour

End Sub


'---------------------------------------------------------------------
Private Function GetResponse(ByVal lPersonId As Long, _
                                ByVal lVisitId As Long, _
                                ByVal nVisitCycleNumber As Integer, _
                                ByVal lCRFPageId As Long, _
                                ByVal nCRFPageCycleNumber As Integer, _
                                ByVal lDataItemId As Long, _
                                ByVal nRepeatNumber As Integer, _
                                ByVal nPass As Integer, _
                                ByRef sUserName As String) As String
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsDoubleData As ADODB.Recordset
Dim sDDResponse As String

    sSQL = "SELECT Response, UserName FROM DoubleData " _
        & " WHERE ClinicalTrialId = " & mlVerTrialId _
        & " AND TrialSite = '" & msVerSite & "'" _
        & " AND PersonId = " & lPersonId _
        & " AND VisitId = " & lVisitId _
        & " AND VisitCycleNumber = " & nVisitCycleNumber _
        & " AND CRFPageId = " & lCRFPageId _
        & " AND CRFPageCycleNumber = " & nCRFPageCycleNumber _
        & " AND DataItemId = " & lDataItemId _
        & " AND RepeatNumber = " & nRepeatNumber _
        & " AND Pass = " & nPass
    Set rsDoubleData = New ADODB.Recordset
    rsDoubleData.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    If rsDoubleData.RecordCount = 0 Then
        sDDResponse = ""
        sUserName = ""
    Else
        If IsNull(rsDoubleData!Response) Then
            sDDResponse = ""
        Else
            sDDResponse = rsDoubleData!Response
        End If
        sUserName = rsDoubleData!UserName
    End If
    
    rsDoubleData.Close
    Set rsDoubleData = Nothing
    
    GetResponse = sDDResponse

End Function


'---------------------------------------------------------------------
Private Function GetNextBatchId() As Long
'---------------------------------------------------------------------
' Returns next available BatchResponsId for table BatchResponseData
'---------------------------------------------------------------------
Dim rsNextBatchId As ADODB.Recordset
Dim sSQL As String
Dim lNewBatchId As Long
         
    sSQL = " Select MAX(BatchResponseId) as MaxBatchId FROM BatchResponseData"
    Set rsNextBatchId = New ADODB.Recordset
    rsNextBatchId.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    If IsNull(rsNextBatchId!MaxBatchId) Then
        lNewBatchId = 1
    Else
        lNewBatchId = rsNextBatchId!MaxBatchId + 1
    End If
    rsNextBatchId.Close
    Set rsNextBatchId = Nothing
    
    GetNextBatchId = lNewBatchId
        
End Function


'---------------------------------------------------------------------
Private Sub GenerateBatchDataEntry(ByVal lPersonId As Long, _
                                    ByVal lVisitId As Long, _
                                    ByVal nVisitCycleNumber As Integer, _
                                    ByVal lCRFPageId As Long, _
                                    ByVal nCRFPageCycleNumber As Integer, _
                                    ByVal lDataItemId As Long, _
                                    ByVal nRepeatNumber As Integer, _
                                    ByVal sResponse As String)
'---------------------------------------------------------------------
'This sub will add the verified response details the BatchResponseData Table
'---------------------------------------------------------------------
Dim sSQL As String
Dim lNextBatchId As Long

    On Error GoTo ErrHandler
    
    lNextBatchId = GetNextBatchId
    
    
'    'if its a numeric question run ConvertLocalNumToStandard over the response
'    lDataType = DataTypeFromId(glSelTrialId, glSelDataItemId)
'    Select Case lDataType
'    Case DataType.IntegerData, DataType.LabTest, DataType.Real
'        sDBResponse = ConvertLocalNumToStandard(sResponse)
'    Case Else
'        sDBResponse = sResponse
'    End Select

    'Add the verified entry to table BatchResponseData
    sSQL = "INSERT INTO BatchResponseData (BatchResponseId, ClinicalTrialId, Site, PersonId, " _
        & "VisitId, VisitCycleNumber, VisitCycleDate, CRFPageID, CRFPageCycleNumber, CRFPageCycleDate, " _
        & "DataItemId, RepeatNumber, Response, UserName) " _
        & "VALUES (" & lNextBatchId & "," & mlVerTrialId & ",'" & msVerSite & "'," & lPersonId & "," _
        & lVisitId & "," & nVisitCycleNumber & ",0," & lCRFPageId & "," & nCRFPageCycleNumber & ",0," _
        & lDataItemId & "," & nRepeatNumber & ",'" & ReplaceQuotes(sResponse) & "','" & goUser.UserName & "')"
    MacroADODBConnection.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "GenerateBatchDataEntry")
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
Private Sub ChangeDDStatusToVerified(ByVal lPersonId As Long, _
                                    ByVal lVisitId As Long, _
                                    ByVal nVisitCycleNumber As Integer, _
                                    ByVal lCRFPageId As Long, _
                                    ByVal nCRFPageCycleNumber As Integer, _
                                    ByVal lDataItemId As Long, _
                                    ByVal nRepeatNumber As Integer)
'---------------------------------------------------------------------
'This subroutine will set the DoubleData.status to eDoubleDataStatus.Verified
'for both the Pass=1 and the Pass=2 entries
'---------------------------------------------------------------------
Dim sSQL As String

    sSQL = "UPDATE DoubleData" _
        & " SET Status = " & eDoubleDataStatus.Verified _
        & " WHERE ClinicalTrialId = " & mlVerTrialId _
        & " AND TrialSite = '" & msVerSite & "'" _
        & " AND PersonId = " & lPersonId _
        & " AND VisitId = " & lVisitId _
        & " AND VisitCycleNumber = " & nVisitCycleNumber _
        & " AND CRFPageId = " & lCRFPageId _
        & " AND CRFPageCycleNumber = " & nCRFPageCycleNumber _
        & " AND DataItemId = " & lDataItemId _
        & " AND RepeatNumber = " & nRepeatNumber
        
    MacroADODBConnection.Execute sSQL

End Sub

'---------------------------------------------------------------------
Private Sub txtResponsePass1_LostFocus(Index As Integer)
'---------------------------------------------------------------------

    'validate and save the currently entered/edited response
    Call SaveUpdateChangedResponse(Index, 1)

End Sub

'---------------------------------------------------------------------
Private Sub txtResponsePass2_LostFocus(Index As Integer)
'---------------------------------------------------------------------

    'validate and save the currently entered/edited response
    Call SaveUpdateChangedResponse(Index, 2)

End Sub

'---------------------------------------------------------------------
Private Sub SaveUpdateChangedResponse(ByVal nIndex As Integer, _
                                        ByVal nPassNumber As Integer)
'---------------------------------------------------------------------
Dim sResponse As String
Dim asTag() As String
Dim sSQL As String
Dim lPersonId As Long
Dim lVisitId As Long
Dim nVisitCycleNumber As Integer
Dim lCRFPageId As Long
Dim nCRFPageCycleNumber As Integer
Dim lDataItemId As Long
Dim nFieldOrder As Integer
Dim nRepeatNumber As Integer
Dim nQGroupFieldOrder As Integer
Dim nDataItemType As Integer
Dim rsDoubleData As ADODB.Recordset

    If nPassNumber = 1 Then
        'extract details from Pass1 textbox
        sResponse = txtResponsePass1(nIndex).Text
        asTag = Split(txtResponsePass1(nIndex).Tag, "|")
    Else
        'extract details from Pass2 textbox
        sResponse = txtResponsePass2(nIndex).Text
        asTag = Split(txtResponsePass2(nIndex).Tag, "|")
    End If
    lPersonId = CLng(asTag(0))
    lVisitId = CLng(asTag(1))
    nVisitCycleNumber = CInt(asTag(2))
    lCRFPageId = CLng(asTag(3))
    nCRFPageCycleNumber = CInt(asTag(4))
    lDataItemId = CLng(asTag(5))
    nFieldOrder = CLng(asTag(6))
    nRepeatNumber = CInt(asTag(7))
    nQGroupFieldOrder = CLng(asTag(8))
    nDataItemType = CInt(asTag(9))
            
    If sResponse <> "" Then
        'Check response for invalid characters and length
        If (InStr(sResponse, "`") > 0) Or (InStr(sResponse, """") > 0) Or (InStr(sResponse, "|") > 0) Or (InStr(sResponse, "~") > 0) Then
            Call DialogError("Question response may not contain double or backwards quotes or the | or ~ characters.", "Invalid Response")
            If nPassNumber = 1 Then
                txtResponsePass1(nIndex).SetFocus
            Else
                txtResponsePass2(nIndex).SetFocus
            End If
            Exit Sub
        End If
        If Len(sResponse) > 255 Then
            Call DialogError("Question response may not be longer than 255 characters.", "Invalid Response")
            If nPassNumber = 1 Then
                txtResponsePass1(nIndex).SetFocus
            Else
                txtResponsePass2(nIndex).SetFocus
            End If
            Exit Sub
        End If
        'Validate the response based on DataItemType
        Select Case nDataItemType
        Case DataType.Text, DataType.Category, DataType.Thesaurus
            'no additional validation required
        Case DataType.IntegerData
            'check for numbers only
            If Not gblnValidString(sResponse, valNumeric) Then
                Call DialogError("This Integer question can only contain numeric characters.", "Invalid Response")
                If nPassNumber = 1 Then
                    txtResponsePass1(nIndex).SetFocus
                Else
                    txtResponsePass2(nIndex).SetFocus
                End If
                Exit Sub
            End If
        Case DataType.Real, DataType.LabTest
            'check for numbers and decimal point
            If Not gblnValidString(sResponse, valNumeric + valDecimalPoint) Then
                Call DialogError("This Real/Lab question can only contain numeric characters and decimal points", "Invalid Response")
                If nPassNumber = 1 Then
                    txtResponsePass1(nIndex).SetFocus
                Else
                    txtResponsePass2(nIndex).SetFocus
                End If
                Exit Sub
            End If
        Case DataType.Date
            'check for numbers and date separators /.:- and space
            If Not gblnValidString(sResponse, valNumeric + valDateSeperators) Then
                Call DialogError("This Date/Time question can only contain numeric characters and date separators .:-/", "Invalid Response")
                If nPassNumber = 1 Then
                    txtResponsePass1(nIndex).SetFocus
                Else
                    txtResponsePass2(nIndex).SetFocus
                End If
                Exit Sub
            End If
        End Select
    End If
    
    'Check for the response already existing or being saved for the first time
    sSQL = "SELECT * FROM DoubleData " _
        & " WHERE ClinicalTrialId = " & mlVerTrialId _
        & " AND TrialSite = '" & msVerSite & "'" _
        & " AND PersonId = " & lPersonId _
        & " AND VisitId = " & lVisitId _
        & " AND VisitCycleNumber = " & nVisitCycleNumber _
        & " AND CRFPageID = " & lCRFPageId _
        & " AND CRFPageCycleNumber = " & nCRFPageCycleNumber _
        & " AND DataItemId = " & lDataItemId _
        & " AND RepeatNumber = " & nRepeatNumber _
        & " AND Pass = " & nPassNumber
    Set rsDoubleData = New ADODB.Recordset
    rsDoubleData.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    If rsDoubleData.RecordCount = 0 Then
        rsDoubleData.AddNew
        rsDoubleData.Fields(0) = mlVerTrialId
        rsDoubleData.Fields(1) = msVerSite
        rsDoubleData.Fields(2) = lPersonId
        rsDoubleData.Fields(3) = lVisitId
        rsDoubleData.Fields(4) = nVisitCycleNumber
        rsDoubleData.Fields(5) = lCRFPageId
        rsDoubleData.Fields(6) = nCRFPageCycleNumber
        rsDoubleData.Fields(7) = lDataItemId
        rsDoubleData.Fields(8) = nFieldOrder
        rsDoubleData.Fields(9) = nRepeatNumber
        rsDoubleData.Fields(10) = nQGroupFieldOrder
        rsDoubleData.Fields(11) = nPassNumber
        rsDoubleData.Fields(12) = sResponse
        rsDoubleData.Fields(13) = goUser.UserName
        rsDoubleData.Fields(14) = SQLStandardNow
        rsDoubleData.Fields(15) = eDoubleDataStatus.Entered
        rsDoubleData.Update
    ElseIf (rsDoubleData.Fields(12) <> sResponse) Or (IsNull(rsDoubleData.Fields(12)) And (sResponse <> "")) Then
        rsDoubleData.Fields(12) = sResponse
        rsDoubleData.Fields(13) = goUser.UserName
        rsDoubleData.Fields(14) = SQLStandardNow
        rsDoubleData.Fields(15) = eDoubleDataStatus.Entered
        rsDoubleData.Update
    End If
   
    rsDoubleData.Close
    Set rsDoubleData = Nothing
    
End Sub

'---------------------------------------------------------------------
Private Sub VereFormChecks()
'---------------------------------------------------------------------

    If Me.WindowState = vbMaximized Then
        Exit Sub
    End If

    'Increase width of form if required
    If picVerEForm.Width > picVer.Width Then
        If picVerEForm.Width + 350 > (Screen.Width) Then
            'Window set to Max width
            picVer.Width = Screen.Width - 300
            frmDataVerification.Width = picVer.Width + 300
            VerFormCentre Me
        Else
            'Window increased in width
            picVer.Width = picVerEForm.Width + 50
            frmDataVerification.Width = picVer.Width + 300
            VerFormCentre Me
        End If
    End If
    
    'Increase height of form if required
    If picVerEForm.Height > picVer.Height Then
        If picVerEForm.Height + 5050 > (Screen.Height - 400) Then
            'Window set to Max height
            picVer.Height = Screen.Height - 5400
            frmDataVerification.Height = picVer.Height + 5000
            VerFormCentre Me
        Else
            'Window increased in height
            picVer.Height = picVerEForm.Height + 50
            frmDataVerification.Height = picVer.Height + 5000
            VerFormCentre Me
        End If
    End If

End Sub

'---------------------------------------------------------------------
Public Sub VereFormScrollbars()
'---------------------------------------------------------------------

    'Assess the need for vertical scrollbar
    If VScroll.Visible = True Then
        VScroll.Max = (picVerEForm.Height - picVer.Height) / 10
        VScroll.LargeChange = CInt(((VScroll.Max / 2) / 10) + 1)
        VScroll.SmallChange = CInt(((VScroll.Max / 10) / 10) + 1)
    Else
        If picVerEForm.Height > picVer.Height Then
            Set VScroll.Container = picVer
            VScroll.Max = (picVerEForm.Height - picVer.Height) / 10
            VScroll.Top = 0
            If picVerEForm.Width > picVer.Width Then
                VScroll.Height = picVer.ScaleHeight - HScroll.Height
            Else
                VScroll.Height = picVer.ScaleHeight
            End If
            VScroll.Left = picVer.ScaleWidth - VScroll.Width
            VScroll.LargeChange = CInt(((VScroll.Max / 2) / 10) + 1)
            VScroll.SmallChange = CInt(((VScroll.Max / 10) / 10) + 1)
            VScroll.Value = 0
            VScroll.Visible = True
        End If
    End If
    
    'Assess the need for horizontal scrollbar
    If picVerEForm.Width > picVer.Width Then
        Set HScroll.Container = picVer
        HScroll.Max = (picVerEForm.Width - picVer.Width) / 10
        HScroll.Top = picVer.ScaleHeight - HScroll.Height
        If VScroll.Visible = True Then
            HScroll.Width = picVer.ScaleWidth - VScroll.Width
        Else
            HScroll.Width = picVer.ScaleWidth
        End If
        HScroll.Left = 0
        HScroll.LargeChange = CInt(((HScroll.Max / 2) / 10) + 1)
        HScroll.SmallChange = CInt(((HScroll.Max / 10) / 10) + 1)
        HScroll.Value = 0
        HScroll.Visible = True
    End If

End Sub

'---------------------------------------------------------------------
Public Sub VerFormCentre(frmForm As Form)
'---------------------------------------------------------------------
'   Centre DDform on screen , ignoring bottom 400 twips for status bar
'---------------------------------------------------------------------

    If frmDataVerification.WindowState = vbNormal Then
        frmDataVerification.Top = (Screen.Height - 400 - frmDataVerification.Height) \ 2
        frmDataVerification.Left = (Screen.Width - frmDataVerification.Width) \ 2
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub VScroll_Change()
'---------------------------------------------------------------------

    picVerEForm.Top = CSng(-VScroll.Value) * 10

End Sub

'---------------------------------------------------------------------
Private Sub VScroll_Scroll()
'---------------------------------------------------------------------

    picVerEForm.Top = CSng(-VScroll.Value) * 10
    
End Sub
