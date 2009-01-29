VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLaboratories 
   Caption         =   "Laboratories"
   ClientHeight    =   4755
   ClientLeft      =   7395
   ClientTop       =   5640
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   4845
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2220
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3540
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwLab 
      Height          =   3375
      Left            =   180
      TabIndex        =   0
      Top             =   720
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
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
   Begin VB.Frame fraLab 
      Height          =   4215
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   4695
      Begin VB.Label lblPleaseSelect 
         Caption         =   "Please select a laboratory for the laboratory test questions on this eForm"
         Height          =   435
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmLaboratories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       frmLaboratories.frm
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Toby Aldridge, September 2000
'   Purpose:    Form for choosing a laboratory in Data Entry
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'
' REVISIONS:
'   NCJ 26/9/00 - Added label saying "Please select..."
'               Added icon
'   NCJ 5/10/00 - LabId -> LabCode
'   TA 26/10/2000: must explicitly make listview visible just before form is shown
'----------------------------------------------------------------------------------------'

Option Explicit

Private msLabCode As String

'----------------------------------------------------------------------------------------'
Public Function Display(sSite As String, ByRef sLabCode As String, _
                        Optional sClinicalTestCode As String = "") As Boolean
'----------------------------------------------------------------------------------------'
' Display list of labs and return selected lab code
' sLabCode will be initially selected lab if non-empty
'   function returns false if cancel clicked or sLabCode unchanged
'   otherwise returns true and sLabCode will be new selected lab
'----------------------------------------------------------------------------------------'
Dim oLabs As clsLabs

    On Error GoTo ErrHandler
    
    msLabCode = ""
    
    FormCentre Me
    
    Me.Icon = frmMenu.Icon
    
    'fill list view
    Set oLabs = New clsLabs
    'restrict by site
    oLabs.Load (sSite)
    Call oLabs.PopulateListView(lvwLab, False)
    
    If lvwLab.ListItems.Count > 0 Then
        If sLabCode = "" Then
            'select first item
            lvwLab.SelectedItem = lvwLab.ListItems(1)
        Else
            'select the lab passed in
            lvwLab.SelectedItem = ListItembyTag(lvwLab, sLabCode)
        End If
    
        If lvwLab.ListItems.Count = 1 Then
            'only one lab - get its id
            'TA 3/10/00: first listitem index is 1
            msLabCode = lvwLab.ListItems(1).Tag
        Else
            DefaultPointerOn
            'TA 26/10/2000: must explicitly make listview visible
            lvwLab.Visible = True
            Me.Show vbModal
            DefaultPointerOff
        End If
    Else
        ' No labs to choose from
        ' Possibly show message here???
    End If
    
    ' If they selected one and it has changed, return the ID
    If msLabCode <> "" And sLabCode <> msLabCode Then
        sLabCode = msLabCode
        Display = True
    Else
        Display = False
    End If

Exit Function
    
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "Display")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Function

'----------------------------------------------------------------------------------------'
Private Sub cmdCancel_Click()
'----------------------------------------------------------------------------------------'

    Unload Me
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub cmdOK_Click()
'----------------------------------------------------------------------------------------'
' unload form and get lab id from sleected listitem
'----------------------------------------------------------------------------------------'
    
    If lvwLab.ListItems.Count > 0 Then
        msLabCode = lvwLab.SelectedItem.Tag
    End If
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenu.Icon
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Form_Resize()
'----------------------------------------------------------------------------------------'

    If Me.Width > 3000 Then
        fraLab.Width = Me.ScaleWidth - 120
        lvwLab.Width = fraLab.Width - 240
        cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 60
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 120
    End If
    
    If Me.Height > 3000 Then
        fraLab.Height = Me.ScaleHeight - cmdOK.Height - 150
        lvwLab.Height = fraLab.Height - lblPleaseSelect.Height - 480
        cmdOK.Top = Me.ScaleHeight - cmdOK.Height - 60
        cmdCancel.Top = cmdOK.Top
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub lvwLab_DblClick()
'----------------------------------------------------------------------------------------'
'unload form
'----------------------------------------------------------------------------------------'

    cmdOK_Click
    
End Sub
