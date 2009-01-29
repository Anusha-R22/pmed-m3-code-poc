VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSDImages 
   Caption         =   "This form stores images only and is never displayed"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imgList 
      Left            =   780
      Top             =   900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSDImages.frx":0000
            Key             =   "Hotlink"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSDImages.frx":031A
            Key             =   "link2"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSDImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       frmSDImages.frm
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   Author:     Nicky Johns, November 2002
'   Purpose:    Stores images for MACRO SD
'               This form is never visible!
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
' REVISIONS:
' NCJ 7 Nov 02 - Initial Development
'----------------------------------------------------------------------------------------'

Option Explicit

' There is no code!


