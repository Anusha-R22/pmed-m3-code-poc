VERSION 5.00
Begin VB.Form frmDistributionResults 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Study Distribution Results"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   3300
      TabIndex        =   1
      Top             =   2280
      Width           =   1125
   End
   Begin VB.TextBox txtDistributeResults 
      Enabled         =   0   'False
      Height          =   2115
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   4395
   End
End
Attribute VB_Name = "frmDistributionResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1998-2002. All Rights Reserved
'   File:       frmDistributionResults.frm
'   Author:     David Hook, 09/08/2002
'   Purpose:    Display the results of the last distribution
'--------------------------------------------------------------------------------
' REVISIONS
' DPH 04/09/2002 - Make form for Distribution & Recall Results
'--------------------------------------------------------------------------------

Option Explicit
Private msDistributeResults As String

'--------------------------------------------------------------------------------
Private Sub cmdOK_Click()
'--------------------------------------------------------------------------------
' Close Form
'--------------------------------------------------------------------------------

    Unload Me

End Sub

'--------------------------------------------------------------------------------
Private Sub Form_Load()
'--------------------------------------------------------------------------------
' Load form setting up listview and distribution details
'--------------------------------------------------------------------------------

    ' set up icon
    Me.Icon = frmMenu.Icon

    ' set up results
    txtDistributeResults.Text = msDistributeResults
    txtDistributeResults.Enabled = True

End Sub

'--------------------------------------------------------------------------------
Public Sub InitialiseMe(sResults As String, bDistribute As Boolean)
'--------------------------------------------------------------------------------
' Initialise form with the distributiion results
'--------------------------------------------------------------------------------
' REVISIONS
' DPH 04/09/2002 - Make form for Distribution & Recall Results
'--------------------------------------------------------------------------------

    msDistributeResults = sResults

    If bDistribute = True Then
        Me.Caption = "Study Distribution Results"
    Else
        Me.Caption = "Recalled Distribution Results"
    End If

End Sub

