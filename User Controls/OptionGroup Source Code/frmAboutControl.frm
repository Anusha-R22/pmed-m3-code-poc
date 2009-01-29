VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About app.title"
   ClientHeight    =   795
   ClientLeft      =   3795
   ClientTop       =   3960
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   795
   ScaleWidth      =   5085
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4020
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Copyright 2002, InferMed Ltd"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label lblAbout 
      Caption         =   "app.title Version app.major.app.minor"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   File:       frmAbout.frm
'   Author:     Zulfiqar Ahmed, September 2001
'   Purpose:    About dialogue for OptionGroup control
'----------------------------------------------------------------------------------
Option Explicit

'----------------------------------------------------------------------------------
Private Sub cmdOK_Click()
'----------------------------------------------------------------------------------
'unload the form if user presses OK button
'----------------------------------------------------------------------------------
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblAbout.Caption = App.Title & " Version " & App.Major & "." & App.Minor
End Sub
