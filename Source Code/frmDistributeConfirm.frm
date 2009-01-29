VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmDistributeConfirm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Study Distribution Confirmation"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   5340
      TabIndex        =   2
      Top             =   3120
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4140
      TabIndex        =   1
      Top             =   3120
      Width           =   1125
   End
   Begin MSComctlLib.ListView lvwConfirmList 
      Height          =   2535
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblConfirm 
      Caption         =   "Please confirm distribution of the following study version(s) :"
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   6375
   End
End
Attribute VB_Name = "frmDistributeConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1998-2002. All Rights Reserved
'   File:       frmDistributeConfirm.frm
'   Author:     David Hook, 09/08/2002
'   Purpose:    Confirm distribution of Study version(s) to remote sites
'--------------------------------------------------------------------------------

Option Explicit
Public mbDistribute As Boolean

'--------------------------------------------------------------------------------
Private Sub cmdCancel_Click()
'--------------------------------------------------------------------------------
' Cancel out of the form
'--------------------------------------------------------------------------------

    mbDistribute = False
    Me.Hide
    
End Sub

'--------------------------------------------------------------------------------
Private Sub cmdOK_Click()
'--------------------------------------------------------------------------------
' Set distribute boolean to true and hide form
'--------------------------------------------------------------------------------

    mbDistribute = True
    Me.Hide
    
End Sub

'--------------------------------------------------------------------------------
Private Sub Form_Load()
'--------------------------------------------------------------------------------
' Load form setting up icon
'--------------------------------------------------------------------------------

    ' set icon
    Me.Icon = frmMenu.Icon
    mbDistribute = False
    lvw_SetAllColWidths lvwConfirmList, LVSCW_AUTOSIZE_USEHEADER
    
End Sub

'--------------------------------------------------------------------------------
Public Sub AddDetailToListView(sSiteCode As String, sStudyCode As String, _
                lVersion As Long, sVersionDescription As String)
'--------------------------------------------------------------------------------
' Add detail to form's listview a row at a time
'--------------------------------------------------------------------------------
Dim oLIDistribution As ListItem
    
    On Error GoTo ErrorHandler

    If lvwConfirmList.ColumnHeaders.Count = 0 Then
        lvwConfirmList.ColumnHeaders.Add , , "Site", 1400
        lvwConfirmList.ColumnHeaders.Add , , "Study", 1400
        lvwConfirmList.ColumnHeaders.Add , , "Version", 800
        lvwConfirmList.ColumnHeaders.Add , , "Description", 2500
    End If
    
    Set oLIDistribution = lvwConfirmList.ListItems.Add(, , sSiteCode)
    oLIDistribution.SubItems(1) = sStudyCode
    oLIDistribution.SubItems(2) = lVersion
    oLIDistribution.SubItems(3) = sVersionDescription

Exit Sub
ErrorHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "GetStudySiteLatestVersion")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
  End Select
End Sub
