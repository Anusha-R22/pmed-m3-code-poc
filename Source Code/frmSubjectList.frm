VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSubjectList 
   BorderStyle     =   0  'None
   Caption         =   "Subject List"
   ClientHeight    =   7275
   ClientLeft      =   2145
   ClientTop       =   1815
   ClientWidth     =   9600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7275
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   8280
      TabIndex        =   6
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Frame fraSearchCriteria 
      Caption         =   "Search Criteria"
      Height          =   855
      Left            =   60
      TabIndex        =   7
      Top             =   0
      Width           =   9435
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   8040
         TabIndex        =   4
         Top             =   270
         Width           =   1215
      End
      Begin VB.TextBox txtSite 
         Height          =   345
         Left            =   4425
         TabIndex        =   2
         Top             =   283
         Width           =   1200
      End
      Begin VB.TextBox txtStudy 
         Height          =   345
         Left            =   2700
         TabIndex        =   1
         Top             =   283
         Width           =   1200
      End
      Begin VB.TextBox txtSubjectId 
         Height          =   345
         Left            =   6633
         MaxLength       =   10
         TabIndex        =   3
         Top             =   283
         Width           =   1140
      End
      Begin VB.TextBox txtLabel 
         Height          =   345
         Left            =   789
         TabIndex        =   0
         Top             =   283
         Width           =   1200
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   11
         Top             =   330
         Width           =   390
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject Id"
         Height          =   195
         Index           =   1
         Left            =   5820
         TabIndex        =   10
         Top             =   330
         Width           =   720
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Site"
         Height          =   195
         Index           =   0
         Left            =   4080
         TabIndex        =   9
         Top             =   330
         Width           =   270
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Study"
         Height          =   195
         Index           =   8
         Left            =   2220
         TabIndex        =   8
         Top             =   330
         Width           =   405
      End
   End
   Begin MSComctlLib.ListView lvwSubjects 
      Height          =   5475
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   9657
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   0
   End
End
Attribute VB_Name = "frmSubjectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       frmOpenSubject.bas
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   Author:     Toby Aldridge, August 2001
'   Purpose:    New Open Subject Form (based on 2.1 version)
'----------------------------------------------------------------------------------------'
' Revisions:
'   TA10/09/2001: haven't done any fiddly are-things-selected-valid
'                  or have-they-typed-in-tilde-characters-etc-stuff yet
'   TA 02/10/2001: form does not unload - only hides
'   TA 03/10/2001: Changes so that sites are filtered by UserName
'   TA 10/10/2001: Onlt alow numbers in subjectid field
'   ASH 8/11/2001: Incremented Ubound of array by 1 to get correct number of patients (RefreshList)
'   TA 01/10/2002: New UI Improvements
'   TA 01/11/2002 - Changed to new icons
'----------------------------------------------------------------------------------------'

Option Explicit

Private Const msKEY_SEPARATOR = "|"

'store connection
Private msConString As String

Public Event SubjectSelected(lStudyId As Long, sSite As String, lSubjectId As Long)

'---------------------------------------------------------------------
Public Sub Display(sCon As String, lTop As Long, lLeft As Long, lHeight As Long, lWidth As Long, _
                        Optional sStudyName As String = "", Optional sSite As String = "", Optional sLabel As String = "")
'---------------------------------------------------------------------

        Load Me
        
        msConString = gsADOConnectString
        
        With Me
            .Top = lTop
            .Left = lLeft
            .Width = lWidth
            .Height = lHeight
        End With
        
        Me.Show vbModeless
        Me.ZOrder
        
        If sStudyName & sSite & sLabel <> "" Then
            'we ahve been passed in some filter info
            txtStudy.Text = sStudyName
            txtSite.Text = sSite
            txtLabel.Text = sLabel
            RefreshList
        End If
        
End Sub
'---------------------------------------------------------------------
Private Sub cmdCancel_Click()
'---------------------------------------------------------------------

    Me.Hide
    'inform everyone that i'm closed
    CloseWinForm wfDataBrowser
End Sub

'---------------------------------------------------------------------
Private Sub cmdOK_Click()
'---------------------------------------------------------------------
Dim sKey As String
Dim lStudyId As Long
Dim sSite As String
Dim lSubjectId As Long
    
    
    sKey = lvwSubjects.SelectedItem.Tag
        
    lStudyId = Split(sKey, msKEY_SEPARATOR)(0)
    sSite = Split(sKey, msKEY_SEPARATOR)(1)
    lSubjectId = Split(sKey, msKEY_SEPARATOR)(2)
    
    RaiseEvent SubjectSelected(lStudyId, sSite, lSubjectId)
    

End Sub


'---------------------------------------------------------------------
Private Sub txtLabel_Change()
'---------------------------------------------------------------------

    CheckSearchText

End Sub

'---------------------------------------------------------------------
Private Sub txtSubjectId_Change()
'---------------------------------------------------------------------
    
    CheckSearchText

End Sub

'---------------------------------------------------------------------
Private Sub txtSubjectId_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
' only allow backspace as integers
'---------------------------------------------------------------------

    If KeyAscii <> Asc(vbBack) And Not gblnValidString(Chr(KeyAscii), valNumeric) Then
        KeyAscii = 0
    End If

End Sub

'---------------------------------------------------------------------
Private Sub txtStudy_Change()
'---------------------------------------------------------------------

    CheckSearchText

End Sub

'---------------------------------------------------------------------
Private Sub txtSite_Change()
'---------------------------------------------------------------------

    CheckSearchText

End Sub

'---------------------------------------------------------------------
Private Sub CheckSearchText()
'---------------------------------------------------------------------

    If txtLabel.Text = "" And txtStudy.Text = "" _
    And txtSite.Text = "" And txtSubjectId.Text = "" Then
        cmdRefresh.Caption = "&Refresh"
    Else
        cmdRefresh.Caption = "&Search"
    End If

End Sub
'---------------------------------------------------------------------
Private Sub Form_Activate()
'---------------------------------------------------------------------

    cmdOK.Enabled = Not (lvwSubjects.SelectedItem Is Nothing)
    
End Sub

'---------------------------------------------------------------------
Private Sub Form_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------

    If KeyAscii = Asc(vbCr) Then
        Select Case Me.ActiveControl.Name
        Case "lvwSubjects"
            cmdOK_Click
        Case "txtLabel", "txtStudy", "txtSite", "txtSubjectId"
            RefreshList
        Case Else
            'do nothing
        End Select
    End If
   
End Sub

'---------------------------------------------------------------------
Private Sub cmdRefresh_Click()
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    RefreshList
    
End Sub

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    Me.Icon = frmMenu.Icon
    
        
    ' Add ColumnHeaders with appropriate widths (5 spaces added to create a bit of padding)
    With lvwSubjects
    
        .ColumnHeaders.Add , , "Status" & Space(5)
        .ColumnHeaders.Add , , "Changed" & Space(5)
        .ColumnHeaders.Add , , "Study" & Space(5)
        .ColumnHeaders.Add , , "Site" & Space(5)
        .ColumnHeaders.Add , , "Id" & Space(5)
        .ColumnHeaders.Add , , "Label" & Space(5)
        
        lvw_SetAllColWidths lvwSubjects, LVSCW_AUTOSIZE_USEHEADER
        
        .Icons = frmImages.imglistStatus
        .SmallIcons = frmImages.imglistStatus
    End With

    cmdOK.Enabled = False

    Me.BackColor = eMACROColour.emcBackGround
    fraSearchCriteria.BackColor = eMACROColour.emcNonWhiteBackGround
    
    fraSearchCriteria.Left = 0
    fraSearchCriteria.BorderStyle = 0
    
    lvwSubjects.Left = 60
    lvwSubjects.BorderStyle = 0
    
   
End Sub

'---------------------------------------------------------------------
Private Sub lvwSubjects_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'---------------------------------------------------------------------
'sort listview
'---------------------------------------------------------------------

    lvw_Sort lvwSubjects, ColumnHeader
    
End Sub

'---------------------------------------------------------------------
Private Sub lvwSubjects_DblClick()
'---------------------------------------------------------------------
'Double click on Listview
'---------------------------------------------------------------------
    
    'ZA 13/06/2002 - check if user has selected an item before calling cmdOK_Click event
    If Not lvwSubjects.SelectedItem Is Nothing Then
        cmdOK_Click
    End If
    
    
End Sub

'---------------------------------------------------------------------
Private Sub lvwSubjects_ItemClick(ByVal Item As MSComctlLib.ListItem)
'---------------------------------------------------------------------

    cmdOK.Enabled = True

End Sub

'---------------------------------------------------------------------
Private Sub Form_Resize()
'---------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    If Me.Height >= fraSearchCriteria.Height + cmdOK.Height + 1920 Then
        lvwSubjects.Height = Me.ScaleHeight - fraSearchCriteria.Height - cmdOK.Height - 240
        cmdOK.Top = lvwSubjects.Top + lvwSubjects.Height + 60
        cmdCancel.Top = cmdOK.Top
    End If
    
    If Me.Width >= cmdOK.Width + cmdCancel.Width + 480 Then
        fraSearchCriteria.Width = Me.ScaleWidth
        lvwSubjects.Width = Me.ScaleWidth - 120
        cmdCancel.Left = lvwSubjects.Left + lvwSubjects.Width - cmdCancel.Width - 120
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 120
    End If
    
ErrHandler:
    
End Sub

Private Sub RefreshList()
'--------------------------------------------
' refresh list when refresh is pressed
' ASH 8/11/2001: Incremented Ubound of array
' by 1 to get correct number of patients
'--------------------------------------------
Dim vData As Variant
Dim lSubjectId As Long
Dim lRow As Long
Dim sKey As String
Dim oItm As ListItem
Dim sIcon As String
Dim oList As UserDataLists

    HourglassOn

  'TA 20/6/2000 SR 3634
    If IsNumeric(txtSubjectId.Text) Then
        lSubjectId = Val(txtSubjectId.Text)
    Else
        lSubjectId = -1
    End If

    
     With goUser.DataLists
            vData = .GetSubjectList((txtLabel.Text), (txtStudy.Text), (txtSite.Text), (lSubjectId))
    End With
       
    If IsNull(vData) Then
        'no subject - warn TA 20/6/2000 SR3634
        Call DialogInformation("No subjects found", "Search Results")
    Else
        lvwSubjects.ListItems.Clear
        For lRow = 0 To UBound(vData, 2)
               
            Set oItm = lvwSubjects.ListItems.Add(, , GetStatusString(ConvertFromNull(vData(eSubjectListCols.SubjectStatus, lRow), vbInteger)))
            
            With oItm
                'the key is Study/Site/Subject
                .Tag = vData(eSubjectListCols.StudyId, lRow) & msKEY_SEPARATOR _
                            & vData(eSubjectListCols.Site, lRow) & msKEY_SEPARATOR _
                            & vData(eSubjectListCols.SubjectId, lRow)
                
                '2nd col is changed flag
                If vData(eSubjectListCols.SubjectStatus, lRow) = 2 Then
                    .ListSubItems.Add , , "New"
                Else
                    .ListSubItems.Add , , ""
                End If
                '3rd col is study
                .ListSubItems.Add , , vData(eSubjectListCols.StudyName, lRow)
                '4th col is site
                .ListSubItems.Add , , vData(eSubjectListCols.Site, lRow)
                '5th col is subjectid
                .ListSubItems.Add , , Right$("000000" & vData(eSubjectListCols.SubjectId, lRow), 6)
                '6th col is subject label
                .ListSubItems.Add , , ConvertFromNull(vData(eSubjectListCols.SubjectLabel, lRow), vbString)
            
                sIcon = GetIconFromStatus(ConvertFromNull(vData(eSubjectListCols.SubjectStatus, lRow), vbInteger))
                If sIcon <> "" Then
                    oItm.SmallIcon = sIcon
                    oItm.Icon = sIcon
                End If
            End With
    
        Next
       
       lvw_SetAllColWidths lvwSubjects, LVSCW_AUTOSIZE_USEHEADER
       'add some to first col for icon
       lvwSubjects.ColumnHeaders(1).Width = lvwSubjects.ColumnHeaders(1).Width + 500
       
        'show count in caption
        'added 1 to UBOUND of array to get correct number of subjects 8/11/2001 ASH
'        fraSubjects.Caption = UBound(vData, 2) + 1 & " subjects found"
'        If lvwSubjects.ListItems.Count = 1 Then
'            'one subject - select them
'            fraSubjects.Caption = "1 subject found"
'        End If
    
    End If
    
    cmdRefresh.Caption = "&Refresh"
    
    cmdOK.Enabled = Not (lvwSubjects.SelectedItem Is Nothing)
    
    HourglassOff
 
End Sub
