VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImages 
   Caption         =   "Image List Container"
   ClientHeight    =   975
   ClientLeft      =   7335
   ClientTop       =   3795
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   975
   ScaleWidth      =   4755
   Visible         =   0   'False
   Begin MSComctlLib.ImageList imglistStatus 
      Left            =   90
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   6
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   30
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":0000
            Key             =   "DM30_ChangeCount1"
            Object.Tag             =   "DM_"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":006B
            Key             =   "DM30_ChangeCount2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":00DB
            Key             =   "DM30_ChangeCount3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":0168
            Key             =   "DM30_RaisedDisc"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":0237
            Key             =   "DM30_Comment"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":05AA
            Key             =   "DM30_Note"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":091C
            Key             =   "DM30_RespondedDisc"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":09E8
            Key             =   "DM30_Frozen"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":0A7F
            Key             =   "DM30_Inform"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":0DF6
            Key             =   "DM30_Locked"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":0EA9
            Key             =   "DM30_Missing"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":1239
            Key             =   "DM30_NA"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":15B1
            Key             =   "DM30_OK"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":1648
            Key             =   "DM30_OKWarning"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":1A09
            Key             =   "DM30_Unobtainable"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":1DAC
            Key             =   "DM30_Warning"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":1E44
            Key             =   "DM30_NewForm"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":1EF6
            Key             =   "DM30_InactiveForm"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":2295
            Key             =   "DM30_Invalid"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":2643
            Key             =   "DM30_QueriedSDV"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":2999
            Key             =   "DM30_DoneSDV"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":2CF3
            Key             =   "DM30_PlannedSDV"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":304A
            Key             =   "DM30_BackEForm"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":3432
            Key             =   "DM30_NextEFormOn"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":381F
            Key             =   "DM30_ChangeCountAll"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":38CD
            Key             =   "DM30_DICTIONARY"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":3C74
            Key             =   "DM30_VDICTIONARY"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":4022
            Key             =   "DM30_CDICTIONARY"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":43D0
            Key             =   "DM30_PDICTIONARY"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":477D
            Key             =   "DM30_XDICTIONARY"
         EndProperty
      EndProperty
   End
   Begin VB.Label CursorHandPoint 
      Caption         =   "CursorHandPoint"
      Height          =   195
      Left            =   900
      MouseIcon       =   "frmImages.frx":4B31
      TabIndex        =   0
      Top             =   180
      Width           =   1395
   End
End
Attribute VB_Name = "frmImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   File:       frmImages.frm
'   Author:     Toby Aldridge, November 2002
'   Purpose:    Simple form to contain status images
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'------------------------------------------------------------------------------------'

Option Explicit


Private Sub Form_Load()

        'add this one manually as we have no gif yet
'    imglistStatus.ListImages.Add , gsVALIDATION_MANDATORY_LABEL, LoadResPicture(gsVALIDATION_MANDATORY_LABEL, vbResIcon)

End Sub
