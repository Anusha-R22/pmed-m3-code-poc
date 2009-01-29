VERSION 5.00
Begin VB.Form frmSiteTransfer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Site Transfer"
   ClientHeight    =   4470
   ClientLeft      =   8415
   ClientTop       =   6270
   ClientWidth     =   3570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   3570
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   60
      TabIndex        =   11
      Top             =   4020
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select the site to transfer subject to"
      Height          =   1035
      Left            =   60
      TabIndex        =   8
      Top             =   2820
      Width           =   3435
      Begin VB.Frame Frame3 
         Caption         =   "New Site (Server/Web)"
         Height          =   675
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3195
         Begin VB.ComboBox cboTransSite 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   240
            Width           =   2955
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select a subject to transfer"
      Height          =   2475
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   3435
      Begin VB.Frame fraSubj 
         Caption         =   "Subject"
         Height          =   675
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   3195
         Begin VB.ComboBox cboSubject 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   240
            Width           =   2955
         End
      End
      Begin VB.Frame fraStudy 
         Caption         =   "Study"
         Height          =   675
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3210
         Begin VB.ComboBox cboStudy 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
            Width           =   2955
         End
      End
      Begin VB.Frame fraSite 
         Caption         =   "Site (Server/Web)"
         Height          =   675
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   3195
         Begin VB.ComboBox cboSite 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   240
            Width           =   2955
         End
      End
   End
   Begin VB.CommandButton cmdTransfer 
      Caption         =   "&Transfer"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   4020
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3600
      Y1              =   2700
      Y2              =   2700
   End
End
Attribute VB_Name = "frmSiteTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
' File:         frmSiteTransfer.frm
' Copyright:    InferMed Ltd. 2003. All Rights Reserved
' Author:       Richard Meinesz, October 2003
' Purpose:      To transfer a subject from one site to another (server/web sites only)
'----------------------------------------------------------------------------------------'
'   Revisions:
'----------------------------------------------------------------------------------------'

Option Explicit

Private mcolWritableSites As Collection

Private mlSelStudyId As Long
Private msSelStudyName As String
Private msSelSite As String
Private mlSelSubjectId As Long
Private msSelTransSite As String
Private msSelSubjectName As String

' The subject list array
Private mvSubjects As Variant

Private mcolSites As Collection

'Public conct that will eventually be moved into basCommon when this is added to SM
Private Const gsSUBJECT_SITE_TRANSFER = "SubjectSiteTransfer"

'--------------------------------------------
Public Sub Display()
'--------------------------------------------
'REM 03/010/03
'Display Site transfer form
'--------------------------------------------
    
    On Error GoTo ErrorLabel
    
    Me.Icon = frmMenu.Icon
    
    'Load all the combos
    Call RefreshSiteTransfer
    
    FormCentre Me
    
    Me.Show vbModal
    
Exit Sub
ErrorLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "frmSiteTransfer", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------
Private Sub RefreshSiteTransfer()
'--------------------------------------------
'REM 06/10/03
'Refresh all the drop downs
'--------------------------------------------

    On Error GoTo ErrorLabel

    cmdTransfer.Enabled = False

    
    If LoadStudies Then
        If LoadSites Then
            If LoadSubjects Then
                Call LoadTransSites
            End If
        End If
    End If
    
    Call TransferOK
    
Exit Sub
ErrorLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmSiteTransfer.RefreshSiteTransfer"
End Sub

'--------------------------------------------
Private Function LoadStudies() As Boolean
'--------------------------------------------
' Populate the Study combo with studies the user has access to
'--------------------------------------------
Dim lRow As Long
Dim vStudies As Variant
Dim oStudy As Study
Dim colStudies As Collection
Dim bLoadStudies As Boolean
    
    On Error GoTo ErrorLabel

    HourglassOn
    
    cboStudy.Clear
    
    Set colStudies = goUser.GetAllStudies
    
    ' Are there any studies?
    If colStudies.Count = 0 Then
        bLoadStudies = False
    Else
        ' Add the studies to the combo
        ' and the study IDs to the ItemData array
        For Each oStudy In colStudies
            cboStudy.AddItem oStudy.StudyName
            cboStudy.ItemData(cboStudy.NewIndex) = oStudy.StudyId
        Next
        cboStudy.ListIndex = 0
        bLoadStudies = True
    End If
 
    LoadStudies = bLoadStudies
 
    HourglassOff
    
Exit Function
ErrorLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmSiteTransfer.LoadStudies"
End Function

'--------------------------------------------
Private Function LoadSites() As Boolean
'--------------------------------------------
' Populate the Sites combo with sites the user has access to
' according to the chosen study
'--------------------------------------------
Dim lRow As Long
Dim vSites As Variant
Dim oSite As Site
Dim bLoadSites As Boolean

    On Error GoTo ErrorLabel

    cboSite.Clear
    msSelSite = ""
    HourglassOn
    
    ' Get all sites
    Set mcolSites = goUser.GetAllSites(cboStudy.ItemData(cboStudy.ListIndex))
    Set mcolWritableSites = New Collection
    
    ' Are there any sites?
    If mcolSites.Count > 0 Then
        For Each oSite In mcolSites
            ' Can't open subjects from Remote sites on the Server
            If Not (goUser.DBIsServer And oSite.SiteLocation = TypeOfInstallation.RemoteSite) Then
                cboSite.AddItem oSite.Site
                mcolWritableSites.Add LCase(oSite.Site), LCase(oSite.Site)
            End If
        Next
    End If
    
    'if there are sites then select the first one
    If cboSite.ListCount > 0 Then
        cboSite.ListIndex = 0
        bLoadSites = True
    Else
        bLoadSites = False
    End If
    
    LoadSites = bLoadSites
    
    HourglassOff
    
Exit Function
ErrorLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmSiteTransfer.LoadSites"
End Function

'--------------------------------------------
Private Sub LoadTransSites()
'--------------------------------------------
'REM 03/10/03
'Load the New Sites combo with all Site for teh selected Study exclusing the selected original Site
'--------------------------------------------
Dim lRow As Long
Dim vSites As Variant
Dim oSite As Site
Dim lCountSites As Long
Dim sSite As String
    
    On Error GoTo ErrorLabel
    
    'clear combo box
    cboTransSite.Clear
    
    lCountSites = 0
    ' Are there any sites?
    If mcolSites.Count > 0 Then
        For Each oSite In mcolSites
            sSite = oSite.Site
            ' Can't otransfer subjects to Remote sites on the Server or to current Subject's site
            If Not (goUser.DBIsServer And oSite.SiteLocation = TypeOfInstallation.RemoteSite) And (sSite <> msSelSite) Then
                cboTransSite.AddItem sSite
                'count the number of sites added to the combo box
                lCountSites = lCountSites + 1
            End If
        Next
    End If
    
    'if there are sites then select the first one
    If lCountSites > 0 Then
        cboTransSite.ListIndex = 0
    End If
    
Exit Sub
ErrorLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmSiteTransfer.LoadTransSites"
End Sub

'--------------------------------------------
Private Function LoadSubjects() As Boolean
'--------------------------------------------
'REM 03/10/03
'Load all the Subjects for the given study and site
'--------------------------------------------
Dim lStudyId As Long
Dim sSite As String
Dim lSubjectRow As Long
Dim sSubjectRow As String
Dim bLoadSubjects As Boolean
    
    On Error GoTo ErrorLabel
    
    cboSubject.Clear

    If GetSelectedSubjects Then
        ' The mvSubjects array will now have been filled in
        For lSubjectRow = 0 To UBound(mvSubjects, 2)
            ' Only take subjects from writable sites
            If CollectionMember(mcolWritableSites, LCase(mvSubjects(eSubjectListCols.Site, lSubjectRow)), False) Then
               
                ' If subject label is available, use it, otherwise use subject ID
                If mvSubjects(eSubjectListCols.SubjectLabel, lSubjectRow) > "" Then
                    cboSubject.AddItem mvSubjects(eSubjectListCols.SubjectLabel, lSubjectRow)
                    cboSubject.ItemData(cboSubject.NewIndex) = mvSubjects(eSubjectListCols.SubjectId, lSubjectRow)
                Else
                    cboSubject.AddItem mvSubjects(eSubjectListCols.SubjectId, lSubjectRow)
                    cboSubject.ItemData(cboSubject.NewIndex) = mvSubjects(eSubjectListCols.SubjectId, lSubjectRow)
                End If

            End If
        Next
    End If
    
    'if there are sites then select the first one
    If cboSubject.ListCount > 0 Then
        cboSubject.ListIndex = 0
        bLoadSubjects = True
    Else
        bLoadSubjects = False
    End If

    LoadSubjects = bLoadSubjects
    
Exit Function
ErrorLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmSiteTransfer.LoadSubjects"
End Function

'--------------------------------------------------------------------
Private Function GetSelectedSubjects() As Boolean
'--------------------------------------------------------------------
' Get the array of subjects into mvSubjects
' according to currently selected Study, Site and Subject
' Returns FALSE if no subjects found
'--------------------------------------------------------------------
Dim sSite As String
Dim lSubjectId As Long
Dim sSubjectLabel As String

    On Error GoTo ErrorLabel
    
    ' The subject label
    sSubjectLabel = ""
    
    If msSelStudyName > "" And msSelSite > "" Then

        sSite = msSelSite

        mvSubjects = goUser.DataLists.GetSubjectList(sSubjectLabel, msSelStudyName, sSite)
        If IsNull(mvSubjects) Then
            ' If the subject label is numeric, try again using it as a subject ID
            If IsNumeric(sSubjectLabel) Then
                lSubjectId = CLng(sSubjectLabel)
                ' Check it's a sensible subject id
                If lSubjectId > 0 And lSubjectId = CDbl(sSubjectLabel) Then
                    mvSubjects = goUser.DataLists.GetSubjectList(, msSelStudyName, sSite, lSubjectId)
                End If
            End If
        End If
        GetSelectedSubjects = Not IsNull(mvSubjects)
    Else
        ' Nothing selected
        GetSelectedSubjects = False
    End If
    
Exit Function
ErrorLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmSiteTransfer.GetSelectedSubjects"
End Function

'--------------------------------------------------------------------
Private Sub cboTransSite_Click()
'--------------------------------------------------------------------
' They clicked on a Site to Transfer to
'--------------------------------------------
    
    If cboSite.ListIndex > -1 Then
        If cboTransSite.List(cboTransSite.ListIndex) <> msSelTransSite Then
            msSelTransSite = cboTransSite.List(cboTransSite.ListIndex)
        End If
    Else
        If msSelTransSite > "" Then
            msSelTransSite = ""
        End If
    End If

    Call TransferOK

End Sub

'--------------------------------------------------------------------
Private Sub cboSite_Click()
'--------------------------------------------------------------------
' They clicked on a Site
'--------------------------------------------
    
    If cboSite.ListIndex > -1 Then
        If cboSite.List(cboSite.ListIndex) <> msSelSite Then
            msSelSite = cboSite.List(cboSite.ListIndex)
            Call LoadSubjects
            Call LoadTransSites
        End If
    Else
        If msSelSite > "" Then
            msSelSite = ""
        End If
    End If
    
    Call TransferOK
    
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

    Call TransferOK
    
End Sub

'--------------------------------------------
Private Sub cboSubject_Click()
'--------------------------------------------
' They clicked on a Subject
'--------------------------------------------

    If cboSubject.ListIndex > -1 Then
        If (cboSubject.ItemData(cboSubject.ListIndex) <> mlSelSubjectId) Or (cboSubject.List(cboSubject.ListIndex) <> msSelSubjectName) Then
            mlSelSubjectId = cboSubject.ItemData(cboSubject.ListIndex)
            msSelSubjectName = cboSubject.List(cboSubject.ListIndex)
        End If
    Else
        If mlSelSubjectId > "" Then
            mlSelSubjectId = ""
        End If
    End If
    
    Call TransferOK
    
End Sub

'--------------------------------------------
Private Sub TransferOK()
'--------------------------------------------
'REM 06/10/03
'Routine to enable combo boxes and Transfer button
'--------------------------------------------

    cboStudy.Enabled = (cboStudy.ListIndex > -1)
    cboSite.Enabled = (cboSite.ListIndex > -1)
    cboSubject.Enabled = (cboSubject.ListIndex > -1)
    cboTransSite.Enabled = (cboTransSite.ListIndex > -1)
    
    'if there are no subjects then clear the Trans Site combo box
    If Not (cboSubject.ListIndex > -1) Then
        cboTransSite.Clear
    End If
    
    cmdTransfer.Enabled = (cboSubject.ListIndex > -1) And (cboTransSite.ListIndex > -1)

End Sub

'--------------------------------------------
Private Sub cmdCancel_Click()
'--------------------------------------------
'exit form
'--------------------------------------------

    Unload Me

End Sub

'--------------------------------------------
Private Sub cmdTransfer_Click()
'--------------------------------------------
'REM 03/10/03
'Transfer a Subject from one site to another
'--------------------------------------------
Dim oSiteTransfer As clsSiteTransfer
Dim sMessage As String
Dim sTransMessage As String
    
    On Error GoTo ErrorLabel
    
    'ask the user is they are sure they want to do the selected transfer
    If DialogQuestion("Transfer subject '" & msSelSubjectName & "' from site '" & msSelSite & "' to site '" & msSelTransSite & "'?") = vbYes Then
        
        HourglassOn
        
        Set oSiteTransfer = New clsSiteTransfer
        
        If oSiteTransfer.PatientSiteTransfer(goUser.CurrentDBConString, goUser.UserName, mlSelStudyId, msSelSite, msSelTransSite, mlSelSubjectId, sMessage) Then
            sTransMessage = "Subject '" & msSelSubjectName & "' has been successfully transferred from site '" & msSelSite & "' to site '" & msSelTransSite & "'"
            Call DialogInformation(sTransMessage)
            Call gLog(gsSUBJECT_SITE_TRANSFER, sTransMessage)
        Else
            sTransMessage = "Unable to transfer subject '" & msSelSubjectName & "' from site '" & msSelSite & "' to site '" & msSelTransSite & "'. " & vbCrLf & sMessage
            Call DialogWarning(sTransMessage)
            Call gLog(gsSUBJECT_SITE_TRANSFER, sTransMessage)
        End If
        
        HourglassOff
        
        'after doing the transfer refresh the form so it shows the Subject in the new site
        Call RefreshSiteTransfer
        
        cmdCancel.Caption = "Close"
        
        Set oSiteTransfer = Nothing
    End If
    
Exit Sub
ErrorLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "SiteTransfer", Err.Source) = Retry Then
        Resume
    End If
End Sub

