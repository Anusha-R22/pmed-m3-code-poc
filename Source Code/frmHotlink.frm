VERSION 5.00
Begin VB.Form frmHotlink 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hotlink Definition"
   ClientHeight    =   3900
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5628
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5628
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraEForm 
      Caption         =   "eForm"
      Height          =   2560
      Left            =   2870
      TabIndex        =   9
      Top             =   800
      Width           =   2715
      Begin VB.CheckBox chkFCycle 
         Caption         =   "Specify eForm cycle"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   630
         Width           =   2295
      End
      Begin VB.OptionButton optFormCycle 
         Caption         =   "Last"
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   17
         Top             =   2160
         Width           =   1035
      End
      Begin VB.OptionButton optFormCycle 
         Caption         =   "First"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   16
         Top             =   1860
         Width           =   1035
      End
      Begin VB.OptionButton optFormCycle 
         Caption         =   "Previous"
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   15
         Top             =   1560
         Width           =   1035
      End
      Begin VB.OptionButton optFormCycle 
         Caption         =   "Next"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   14
         Top             =   1260
         Width           =   1035
      End
      Begin VB.OptionButton optFormCycle 
         Caption         =   "This"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   6
         Top             =   960
         Width           =   1035
      End
      Begin VB.ComboBox cboForms 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   270
         Width           =   2355
      End
   End
   Begin VB.Frame fraVisit 
      Caption         =   "Visit"
      Height          =   2560
      Left            =   60
      TabIndex        =   8
      Top             =   800
      Width           =   2715
      Begin VB.CheckBox chkVCycle 
         Caption         =   "Specify visit cycle"
         Height          =   255
         Left            =   180
         TabIndex        =   18
         Top             =   630
         Width           =   2355
      End
      Begin VB.OptionButton optVisitCycle 
         Caption         =   "Last"
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   13
         Top             =   2160
         Width           =   1035
      End
      Begin VB.OptionButton optVisitCycle 
         Caption         =   "First"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   12
         Top             =   1860
         Width           =   1035
      End
      Begin VB.OptionButton optVisitCycle 
         Caption         =   "Previous"
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   11
         Top             =   1560
         Width           =   1035
      End
      Begin VB.OptionButton optVisitCycle 
         Caption         =   "Next"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   10
         Top             =   1260
         Width           =   1035
      End
      Begin VB.OptionButton optVisitCycle 
         Caption         =   "This"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   5
         Top             =   960
         Width           =   1035
      End
      Begin VB.ComboBox cboVisits 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   2355
      End
   End
   Begin VB.Frame fraCaption 
      Caption         =   "Caption"
      Height          =   700
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   5505
      Begin VB.TextBox txtCaption 
         Height          =   315
         Left            =   90
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   5315
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   3420
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   3420
      Width           =   1215
   End
End
Attribute VB_Name = "frmHotlink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'----------------------------------------------------------------------------------------'
'   File:       frmHotlink.frm
'   Copyright:  InferMed Ltd. 2002-2006. All Rights Reserved
'   Author:     Nicky Johns, November 2002
'   Purpose:    Allows Study Designer to edit a Hotlink element
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
' REVISIONS:
' NCJ 7-8 Nov 02 - Initial Development
' NCJ 22 Jun 04 - Bug 2300 - Warn when editing Hotlink for non-existent visit/eForm
' NCJ 23 Jun 04 - Don't include visit in combo if it only contains a visit eform
' NCJ 14 Jun 06 - Include study access mode
'----------------------------------------------------------------------------------------'

Option Explicit

Private mbClickedOK As Boolean
' The selected visit/form/cycles
Private msSelVisit As String
Private mlVisitId As Long
Private msSelForm As String
Private msSelVCycle As String
Private msSelFCycle As String
Private msCaption As String
Private mbVisitCycle As Boolean
Private mbFormCycle As Boolean

Private mlTrialID As Long

Private mbChanged As Boolean

Private mbDefunctTarget As Boolean

Private mbCanEdit As Boolean

' Default visit and form
Private Const msVISIT = "visit"
Private Const msFORM = "form"

'----------------------------------------------------------------------------------------'
Public Function Display(lTrialId As Long, _
                        ByRef sCaption As String, ByRef sHotlink As String, _
                        bEdit As Boolean) As Boolean
'----------------------------------------------------------------------------------------'
' Display the Edit Hotlink form and return the new Caption and Hotlink
' Returns TRUE if user clicked OK, or FALSE if user clicked Cancel
' NCJ 14 Jun 06 - Consider editability
'----------------------------------------------------------------------------------------'

    mlTrialID = lTrialId
    mbCanEdit = bEdit
    
    FormCentre Me
    
    mbClickedOK = False
    ' Store whether there's a non-existent target
    mbDefunctTarget = False
    
    ' Initialise variables
    msSelForm = ""
    mbFormCycle = False
    mbVisitCycle = False
    msSelVisit = ""
    txtCaption.Text = sCaption
    
    Call PopulateFields(sHotlink)
    
    ' Set the changed flag if there was a non-existent target
    mbChanged = mbDefunctTarget
    
    Call EnableOK
    Call EnableForEditing(bEdit)
    
    Me.Show vbModal
    
    ' Get the caption and hotlink
    sCaption = msCaption
    sHotlink = PackHotlink(msSelVisit, msSelVCycle, msSelForm, msSelFCycle)
    
    Display = mbClickedOK And sHotlink > ""

End Function

'----------------------------------------------------------------------------------------'
Private Sub cboForms_Click()
'----------------------------------------------------------------------------------------'
' Click in Forms combo
' Store the selected form
'----------------------------------------------------------------------------------------'

    If cboVisits.ListIndex > -1 Then
        msSelForm = cboForms.Text
        mbChanged = True
    End If
    Call EnableOK

End Sub

'----------------------------------------------------------------------------------------'
Private Sub cboVisits_Click()
'----------------------------------------------------------------------------------------'
' Click in Visits combo
' Store selected visit and refresh the eForms combo according to this visit
'----------------------------------------------------------------------------------------'

    If cboVisits.ListIndex > -1 Then
        msSelVisit = cboVisits.Text
        mbChanged = True
        If msSelVisit <> msVISIT Then
            mlVisitId = cboVisits.ItemData(cboVisits.ListIndex)
        Else
            mlVisitId = 0
        End If
        Call RefreshFormsCombo(msSelVisit)
    End If
    Call EnableOK
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub chkFCycle_Click()
'----------------------------------------------------------------------------------------'
' Specify eForm cycle
' Enable or disable the cycle options
'----------------------------------------------------------------------------------------'
Dim nIndex As Integer

    mbFormCycle = (chkFCycle.Value = vbChecked)
    Call EnableFCycleOpts(mbFormCycle)
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub EnableFCycleOpts(bEnable As Boolean)
'----------------------------------------------------------------------------------------'
' NCJ 14 Jun 06 - Enable or disable the Form Cycle options
'----------------------------------------------------------------------------------------'
Dim nIndex As Integer

    For nIndex = 0 To optFormCycle.Count - 1
        optFormCycle(nIndex).Enabled = bEnable
    Next
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub chkVCycle_Click()
'----------------------------------------------------------------------------------------'
' Specify Visit cycle
' Enable or disable the cycle options
'----------------------------------------------------------------------------------------'

    mbVisitCycle = (chkVCycle.Value = vbChecked)
    Call EnableVCycleOpts(mbVisitCycle)

End Sub

'----------------------------------------------------------------------------------------'
Private Sub EnableVCycleOpts(bEnable As Boolean)
'----------------------------------------------------------------------------------------'
' NCJ 14 Jun 06 - Enable or disable the Visit Cycle options
'----------------------------------------------------------------------------------------'
Dim nIndex As Integer

    For nIndex = 0 To optVisitCycle.Count - 1
        optVisitCycle(nIndex).Enabled = bEnable
    Next

End Sub

'----------------------------------------------------------------------------------------'
Private Sub cmdCancel_Click()
'----------------------------------------------------------------------------------------'
' User clicks Cancel
'----------------------------------------------------------------------------------------'
    
    mbClickedOK = False
    Unload Me       ' Cancel confirmation gets done in Form_Unload
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub cmdOK_Click()
'----------------------------------------------------------------------------------------'
' User clicks OK
'----------------------------------------------------------------------------------------'

    mbClickedOK = True
    Unload Me

End Sub

'----------------------------------------------------------------------------------------'
Private Sub Form_Load()
'----------------------------------------------------------------------------------------'

    Me.Icon = frmMenu.Icon
    txtCaption.MaxLength = 255

End Sub

'----------------------------------------------------------------------------------------'
Private Sub PopulateFields(sHotlink As String)
'----------------------------------------------------------------------------------------'
' Refresh the visit and eForm combos
'----------------------------------------------------------------------------------------'
Dim sVisit As String
Dim sForm As String

    ' Set visit/form/cycles according to unpacked Hotlink
    Call UnpackHotlink(sHotlink, sVisit, msSelVCycle, sForm, msSelFCycle)
    Call SetVCycleOption(msSelVCycle)
    Call SetFCycleOption(msSelFCycle)
    
    Call RefreshVisitsCombo(sVisit)
    Call SetComboSelection(sForm, cboForms)
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub RefreshVisitsCombo(sSelVisit As String)
'----------------------------------------------------------------------------------------'
' Refresh the visit combo and select the given visit
' NCJ 23 Jun 04 - Exclude visits that ONLY contain a visit eForm
'----------------------------------------------------------------------------------------'
Dim tblVisits As clsDataTable
Dim sSQL As String

    On Error GoTo ErrLabel
    
    ' Only include visits that contain eForms
    ' NCJ 23 Jun 04 - not counting visit eForms
    sSQL = "SELECT DISTINCT StudyVisit.VisitCode, StudyVisit.VisitId FROM StudyVisit, StudyVisitCRFPage" _
                & " WHERE StudyVisit.ClinicalTrialId = " & mlTrialID _
                & " AND StudyVisit.ClinicalTrialId = StudyVisitCRFPage.ClinicalTrialId" _
                & " AND StudyVisit.VisitId = StudyVisitCRFPage.VisitId" _
                & " AND StudyVisitCRFPage.eFormUse = " & eEFormUse.User

    Set tblVisits = TableFromSQL(sSQL)
    
    ' Add in "visit" as first item
    If tblVisits.Rows > 0 Then
        tblVisits.Insert RecordBuild(msVISIT, "0"), 1
    Else
        tblVisits.Add RecordBuild(msVISIT, "0")
    End If
    Call TabletoCombo(cboVisits, tblVisits)
    
    ' Select the chosen visit
    Call SetComboSelection(sSelVisit, cboVisits)
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmHotlink.RefreshVisitsCombo"

End Sub

'----------------------------------------------------------------------------------------'
Private Sub RefreshFormsCombo(sVisit As String)
'----------------------------------------------------------------------------------------'
' Refresh the eForm combo according to selected visit
' If selected visit is "visit", populate with all eForms
'----------------------------------------------------------------------------------------'
Dim tblForms As clsDataTable
Dim sSQL As String

    On Error GoTo ErrLabel
    
    sSQL = "SELECT DISTINCT CRFPage.CRFPageCode, CRFPage.CRFPageId FROM CRFPage, StudyVisitCRFPage" _
                & " WHERE CRFPage.ClinicalTrialId = StudyVisitCRFPage.ClinicalTrialId" _
                & " AND CRFPage.ClinicalTrialId = " & mlTrialID
    
    If sVisit <> msVISIT Then
        ' filter on visit, excluding visit eForms for this visit
        sSQL = sSQL & " AND StudyVisitCRFPage.VisitId = " & mlVisitId _
                & " AND StudyVisitCRFPage.CRFPageId =  CRFPage.CRFPageId " _
                & " AND StudyVisitCRFPage.eFormUse = " & eEFormUse.User
    End If
    
    Set tblForms = TableFromSQL(sSQL)
    
    If tblForms.Rows > 0 Then
        tblForms.Insert RecordBuild(msFORM, "0"), 1
    Else
        tblForms.Add RecordBuild(msFORM, "0")
    End If
    Call TabletoCombo(cboForms, tblForms)
    ' Select the first item in the list
    ListCtrl_Pick cboForms, 0

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmHotlink.RefreshFormsCombo"

End Sub

'----------------------------------------------------------------------------------------'
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'

End Sub

'----------------------------------------------------------------------------------------'
Private Sub Form_Unload(Cancel As Integer)
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'
Dim mbWantToClose As Boolean

    Cancel = 0
    mbWantToClose = True
    
    ' NCJ 22 Jun 04 - Check that the OK button is actually enabled
    If mbChanged And Not mbClickedOK And cmdOK.Enabled Then
        mbWantToClose = (DialogQuestion("Are you sure you want to cancel the changes to this hotlink?") = vbYes)
    End If
    
    If Not mbWantToClose Then
        ' Cancel the Unload
        Cancel = 1
    End If

End Sub

'----------------------------------------------------------------------------------------'
Private Sub optFormCycle_Click(Index As Integer)
'----------------------------------------------------------------------------------------'
' Click on Form Cycle option button
'----------------------------------------------------------------------------------------'
    
    msSelFCycle = optFormCycle(Index).Caption
    mbChanged = True

End Sub

'----------------------------------------------------------------------------------------'
Private Sub optVisitCycle_Click(Index As Integer)
'----------------------------------------------------------------------------------------'
' Click on Visit Cycle option button
'----------------------------------------------------------------------------------------'
    
    msSelVCycle = optVisitCycle(Index).Caption
    mbChanged = True

End Sub

'----------------------------------------------------------------------------------------'
Private Sub txtCaption_Change()
'----------------------------------------------------------------------------------------'
' They changed the caption
'----------------------------------------------------------------------------------------'
Dim sCaption As String
Dim nSelPoint As Integer
    
    sCaption = Trim(txtCaption.Text)
    mbChanged = True
    If gblnValidString(sCaption, valOnlySingleQuotes) Then
'        txtCaption.Tag = sCaption
        msCaption = sCaption
    Else
        ' Remember where they are
        nSelPoint = txtCaption.SelStart
        ' Reset to what was there before
'        sCaption = txtCaption.Tag
        txtCaption.Text = msCaption
        ' Reset the insertion point
        If nSelPoint < Len(msCaption) And nSelPoint > 0 Then
            nSelPoint = nSelPoint - 1
        End If
        txtCaption.SelStart = nSelPoint
    End If
    
    Call EnableOK

End Sub

'----------------------------------------------------------------------------------------'
Private Sub UnpackHotlink(sHotlink As String, _
                ByRef sVisit As String, ByRef sVCycle As String, _
                ByRef sForm As String, ByRef sFCycle As String)
'----------------------------------------------------------------------------------------'
' Hotlink is in format Visit(VCycle):Form(FCycle), or empty string (for new Hotlink)
' Unwrap it into its component parts
' (This may get moved to an AREZZO call... )
'----------------------------------------------------------------------------------------'
Dim vData As Variant
Dim vVisit As Variant
Dim vForm As Variant

    ' Initialise cycles to empty strings
    sVCycle = ""
    sFCycle = ""
    
    If sHotlink = "" Then
        sVisit = msVISIT
        sForm = msFORM
    Else
        ' Separate the visit and eForm
        vData = Split(sHotlink, ":")
        
        ' Does visit include a cycle?
        If InStr(1, vData(0), "(") > 0 Then
            ' Pick off the visit & form before the left bracket
            vVisit = Split(vData(0), "(")
            sVisit = vVisit(0)
            ' And take off the right bracket from the cycle
            sVCycle = Left(vVisit(1), Len(vVisit(1)) - 1)
        Else
            ' It's just a visit code
            sVisit = vData(0)
        End If
        
        ' Does form include a cycle?
        If InStr(1, vData(1), "(") > 0 Then
            vForm = Split(vData(1), "(")
            sForm = vForm(0)
            ' And take off the right bracket from the cycle
            sFCycle = Left(vForm(1), Len(vForm(1)) - 1)
         Else
            ' It's just a form code
            sForm = vData(1)
        End If
   End If
    
End Sub

'----------------------------------------------------------------------------------------'
Private Function PackHotlink(sVisit As String, sVCycle As String, _
                            sForm As String, sFCycle As String) As String
'----------------------------------------------------------------------------------------'
' Create the hotlink string from this visit & eForm
' Hotlink is in format Visit(VCycle): Form(FCycle)
' (This may get moved to an AREZZO call... )
'----------------------------------------------------------------------------------------'
Dim sHotlink As String

    ' NCJ 22 Jun 04 - Visit or Form may be blank if they've deleted the visit or form!
    If sVisit > "" And sForm > "" Then
        sHotlink = sVisit
        ' Did they select a visit cycle?
        If mbVisitCycle And msSelVCycle > "" Then
            sHotlink = sHotlink & "(" & sVCycle & ")"
        End If
        
        sHotlink = sHotlink & ":" & sForm
        ' Did they select a form cycle?
        If mbFormCycle And msSelFCycle > "" Then
            sHotlink = sHotlink & "(" & sFCycle & ")"
        End If
    End If
    
    PackHotlink = sHotlink

End Function

'----------------------------------------------------------------------------------------'
Private Sub SetVCycleOption(sVCycle As String)
'----------------------------------------------------------------------------------------'
' Set the option button corresponding to this visit cycle
' If sVCycle is "", uncheck the visit cycle check box
'----------------------------------------------------------------------------------------'
Dim nIndex As Integer

    chkVCycle.Value = vbUnchecked
    For nIndex = 0 To optVisitCycle.Count - 1
        If LCase(sVCycle) = LCase(optVisitCycle(nIndex).Caption) Then
            optVisitCycle(nIndex).Value = True
            ' Check the VCycle box if we have a cycle
            chkVCycle.Value = vbChecked
        Else
            optVisitCycle(nIndex).Value = False
        End If
    Next
    ' Ensure updating of cycle buttons
    Call chkVCycle_Click
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub SetFCycleOption(sFCycle As String)
'----------------------------------------------------------------------------------------'
' Set the option button corresponding to this eform cycle
' If sFCycle is "", uncheck the form cycle check box
'----------------------------------------------------------------------------------------'
Dim nIndex As Integer

    chkFCycle.Value = vbUnchecked
    For nIndex = 0 To optFormCycle.Count - 1
        If LCase(sFCycle) = LCase(optFormCycle(nIndex).Caption) Then
            optFormCycle(nIndex).Value = True
            chkFCycle.Value = vbChecked
        Else
            optFormCycle(nIndex).Value = False
        End If
    Next
    ' Ensure updating of cycle buttons
    Call chkFCycle_Click
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub SetComboSelection(sVF As String, cboBox As ComboBox)
'----------------------------------------------------------------------------------------'
' Set the combobox selection
' NCJ 22 Jun 04 - Detect when they've got something that they can't have any more (Bug 2300)
'----------------------------------------------------------------------------------------'
Dim i As Integer
Dim bChosen As Boolean
Dim sMsg As String

    bChosen = False
    For i = 0 To cboBox.ListCount - 1
        If LCase(cboBox.List(i)) = LCase(sVF) Then
            cboBox.ListIndex = i
            bChosen = True
            Exit For
        End If
    Next
    
    ' NCJ 22 Jun 04 - Warn of defunct targets
    If Not bChosen And sVF > "" Then
        sMsg = "The target "
        If cboBox.Name = "cboForms" Then
            sMsg = sMsg & "eForm '"
        Else
            sMsg = sMsg & "visit '"
        End If
        Call DialogWarning(sMsg & sVF & "' is no longer available for this Hotlink")
        ' Store that we had a "lost" target
        mbDefunctTarget = True
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub EnableOK()
'----------------------------------------------------------------------------------------'
' Enable or disable the OK button
'----------------------------------------------------------------------------------------'

    cmdOK.Enabled = False
    
    If Not mbCanEdit Then Exit Sub      ' NCJ 14 Jun 06
    
    If Trim(txtCaption.Text) = "" Then Exit Sub
    
    If cboForms.Text = "" Then Exit Sub
    
    If cboVisits.Text = "" Then Exit Sub
    
    ' If we get here we're ok!
    cmdOK.Enabled = True
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub EnableForEditing(bEdit As Boolean)
'----------------------------------------------------------------------------------------'
' NCJ 14 Jun 06 - Enable or disable fields according to whether the user can edit or not
'----------------------------------------------------------------------------------------'

    txtCaption.Enabled = bEdit
    chkFCycle.Enabled = bEdit
    chkVCycle.Enabled = bEdit
    Call EnableFCycleOpts(bEdit)
    Call EnableVCycleOpts(bEdit)
    cboForms.Enabled = bEdit
    cboVisits.Enabled = bEdit
    ' NB OK button is dealt with separately
    
End Sub
