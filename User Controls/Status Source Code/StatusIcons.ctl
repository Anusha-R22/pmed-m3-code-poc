VERSION 5.00
Begin VB.UserControl StatusIcons 
   BackStyle       =   0  'Transparent
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4845
   ScaleHeight     =   855
   ScaleWidth      =   4845
   ToolboxBitmap   =   "StatusIcons.ctx":0000
   Begin VB.Label lblMain 
      Caption         =   "Label1"
      Height          =   195
      Left            =   420
      TabIndex        =   0
      Top             =   240
      Width           =   4200
   End
   Begin VB.Image imgMain 
      Height          =   240
      Left            =   60
      Top             =   180
      Width           =   240
   End
End
Attribute VB_Name = "StatusIcons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   File:       StatisFunction.ctl
'   Author:     Zulfiqar Ahmed, October 2001
'   Purpose:    ActiveX control to display a various status icons
'               to be used in Macro 2.2 and above versions.
'
'-----------------------------------------------------------------------------------
Option Explicit

'---------------------------------------------------------------------
Public Function About() As Variant
Attribute About.VB_UserMemId = -552
'---------------------------------------------------------------------
'Display About dialogue for this control
'---------------------------------------------------------------------
    Load frmAbout
    frmAbout.Show vbModal
End Function

'-----------------------------------------------------------------------
Public Property Let Caption(ByVal strCaption As String)
'-----------------------------------------------------------------------
'Use this property to store the caption value for this control
'-----------------------------------------------------------------------
    lblMain.Caption = strCaption
End Property

'-----------------------------------------------------------------------
Public Property Get Caption() As String
'-----------------------------------------------------------------------
'Use this property to retrieve the caption value for this control
'-----------------------------------------------------------------------
    Caption = lblMain.Caption
End Property

'-----------------------------------------------------------------------
Public Property Let Picture(ByVal RES As String)
'-----------------------------------------------------------------------
'Use this property to store the Picture property to be used by imgMain
'-----------------------------------------------------------------------
    imgMain.Picture = frmImages.imglistStatus.ListImages(RES).Picture
'    Debug.Print RES & " - " & imgMain.Picture.Height '/ Screen.TwipsPerPixelY
    UserControl.ScaleHeight = imgMain.Height
End Property

'-----------------------------------------------------------------------
Private Sub UserControl_Resize()
'-----------------------------------------------------------------------
'set up various properties for lblMain and imgMain here
'-----------------------------------------------------------------------
    On Error Resume Next
    lblMain.Top = (UserControl.ScaleHeight - lblMain.Height) / 2
    imgMain.Top = lblMain.Top
    
    'TA 18/02/2003: bit of a hack to get the alignment right for small icons
    Select Case imgMain.Picture.Height / Screen.TwipsPerPixelY
    Case Is < 10
         imgMain.Top = imgMain.Top + 30
    Case Is < 20
         imgMain.Top = imgMain.Top + 20
    End Select
    '**
    
    lblMain.Left = imgMain.Width + 90
    lblMain.Width = UserControl.Width - lblMain.Left
End Sub

'-----------------------------------------------------------------------
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'-----------------------------------------------------------------------
'Caption and Picture properties are being read by the cotrol
'-----------------------------------------------------------------------
    lblMain.Caption = PropBag.ReadProperty("Caption")
    imgMain.Picture = PropBag.ReadProperty("Picture")
End Sub

'-----------------------------------------------------------------------
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'-----------------------------------------------------------------------
' The control writes the Caption and Picture properties for this control
'-----------------------------------------------------------------------
    PropBag.WriteProperty "Caption", lblMain.Caption
    PropBag.WriteProperty "Picture", imgMain.Picture
End Sub
