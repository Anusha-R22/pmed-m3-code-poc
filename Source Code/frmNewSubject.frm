VERSION 5.00
Begin VB.Form frmNewSubject 
   BorderStyle     =   0  'None
   Caption         =   "New Subject"
   ClientHeight    =   3060
   ClientLeft      =   8910
   ClientTop       =   4605
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4590
      TabIndex        =   3
      Top             =   2550
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3270
      TabIndex        =   2
      Top             =   2550
      Width           =   1215
   End
   Begin VB.Frame fraSite 
      Caption         =   "Site"
      Height          =   1575
      Left            =   3060
      TabIndex        =   1
      Top             =   900
      Width           =   2775
      Begin VB.ListBox lstSite 
         Height          =   1230
         Left            =   60
         TabIndex        =   5
         Top             =   240
         Width           =   2595
      End
   End
   Begin VB.Frame fraStudy 
      Caption         =   "Study"
      Height          =   1575
      Left            =   60
      TabIndex        =   0
      Top             =   900
      Width           =   2910
      Begin VB.ListBox lstStudy 
         Height          =   1230
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2595
      End
   End
End
Attribute VB_Name = "frmNewSubject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------
' File: frmNewSubject.frm
' Copyright: InferMed 2001 All Rights Reserved
' Author: Toby Aldridge, InferMed, Aug 2001
' Purpose: Form to allow user to choose a New subject
'----------------------------------------------------
'Revisions
'   TA 03/10/2001: Changes so that sites are filtered by UserName
'   TA 01/10/2002: New UI Improvements
'   TA 18/03/2003: Changed combos to listboxes

Option Explicit

'user
Private moUser As MACROUser

'store selected details
Private mlStudyId As Long
Private msSite As String

Public Event Selected(lStudyId As Long, sSite As String)

'--------------------------------------------
Public Sub Display(oUser As MACROUser, lStudyId As Long, sSite As String, _
                        lTop As Long, lLeft As Long, lHeight As Long, lWidth As Long)
'--------------------------------------------
' Display open subject form
'Input:
    'oUser - logged in user
    'lStudyId - preselected study id, 0 for not preselected
    'sSite - preselected site, "" for not preselected
'Output:
    'function - OK pressed?
    'lStudyId - selected study id
    'sSite - selected site
'REVISIONS:
' REM 17/01/02 - Add the form icon, bug fix 2.2.7 No.8
'--------------------------------------------

    On Error GoTo ErrLabel
    
    Set moUser = oUser
    
    If LoadStudies Then
        HourglassOn
        'REM 17/01/02 - added icon to form
        With Me
            .Top = lTop
            .Left = lLeft
            .Height = lHeight
            .Width = lWidth
        End With
        
        Me.Icon = frmMenu.Icon
        Me.BackColor = eMACROColour.emcBackGround
        'lblTitle.BackColor = eMACROColour.emcTitlebar
        fraStudy.BackColor = eMACROColour.emcBackGround
        fraSite.BackColor = eMACROColour.emcBackGround
        
        
        If lStudyId = 0 Then
            ' Select first in each list
            On Error Resume Next
            lstStudy.ListIndex = 0
            lstSite.ListIndex = 0
            On Error GoTo ErrLabel
        Else
            ' Preselect the default ones
            ListCtrl_Pick lstStudy, lStudyId
            If sSite <> "" Then
                lstSite.Text = sSite
            End If
        End If
        
        HourglassOff
        HourglassSuspend
        Me.Show vbModeless
        Me.ZOrder
        HourglassResume
    
    Else
        'TA 09/10/2000: Clearer wording
        DialogInformation "There are no studies currently available to you"
    End If
        
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "Display", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------
Public Function LoadStudies() As Boolean
'--------------------------------------------

'--------------------------------------------
Dim lRow As Long
Dim vStudies As Variant
Dim oStudy As Study
Dim colStudies As Collection


    HourglassOn
    
    Set colStudies = moUser.GetNewSubjectStudies
    
    ' Are there any studies?
    If colStudies.Count = 0 Then
        LoadStudies = False
    Else
        lstStudy.Clear
        ' Add the studies to the combo
        ' and the study IDs to the ItemData array
        For Each oStudy In colStudies
            lstStudy.AddItem oStudy.StudyName
            lstStudy.ItemData(lstStudy.NewIndex) = oStudy.StudyId
        Next
        LoadStudies = True
    End If
 
    HourglassOff

   
End Function

'--------------------------------------------
Public Function LoadSites() As Boolean
'--------------------------------------------

'--------------------------------------------
Dim lRow As Long
Dim vSites As Variant

Dim colSites As Collection
Dim oSite As Site

    HourglassOn
    Set colSites = moUser.GetNewSubjectSites(lstStudy.ItemData(lstStudy.ListIndex))


    ' Are there any sites?
    If colSites.Count = 0 Then
        LoadSites = False
        HourglassOff
        Exit Function
    End If
    
    lstSite.Clear
    
    For Each oSite In colSites
        lstSite.AddItem oSite.Site
    Next
    
    LoadSites = True
    
    HourglassOff
    
End Function

'--------------------------------------------
Private Sub lstStudy_Click()
'--------------------------------------------
' They clicked on a Study
'--------------------------------------------

    If Not LoadSites Then
        lstSite.Clear
        cmdOK.Enabled = False
        DialogInformation "There are no available sites for this study"
    Else
        cmdOK.Enabled = True
        
        lstSite.ListIndex = 0
    End If

End Sub

'--------------------------------------------
Private Sub cmdCancel_Click()
'--------------------------------------------
' Cancel without choosing a site/study
'--------------------------------------------

    Unload Me
    
End Sub

'--------------------------------------------
Private Sub cmdOK_Click()
'--------------------------------------------
' Pick up the values currently selected
'--------------------------------------------
    
    ' Any study chosen?
    If lstStudy.ListIndex > -1 Then
        mlStudyId = lstStudy.ItemData(lstStudy.ListIndex)
        ' Any site chosen?
        If lstSite.ListIndex > -1 Then
            msSite = lstSite.Text
        Else
            ' No site chosen
        End If
    Else
        ' No study chosen
    End If
    
    If (msSite > "" And mlStudyId > 0) Then
        RaiseEvent Selected(mlStudyId, msSite)
    End If
    
    
End Sub


Private Sub Form_Resize()
        fraStudy.Left = (Me.ScaleWidth - (fraStudy.Width + fraSite.Width + 120)) / 2
        fraSite.Left = fraStudy.Left + fraStudy.Width + 120
        'lblTitle.Left = fraStudy.Left
        
        cmdCancel.Left = fraSite.Left + fraSite.Width - cmdCancel.Width
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 60
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'inform everyone that i'm closed
    CloseWinForm wfNewSubject
End Sub
