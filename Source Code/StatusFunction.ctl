VERSION 5.00
Begin VB.UserControl StatusFunction 
   BackStyle       =   0  'Transparent
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   615
   ScaleWidth      =   4800
   ToolboxBitmap   =   "StatusFunction.ctx":0000
   Begin VB.Label lblMain 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   420
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   2535
   End
End
Attribute VB_Name = "StatusFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   File:       StatisFunction.ctl
'   Author:     Zulfiqar Ahmed, October 2001
'   Purpose:    ActiveX control to display a various status information
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
Private Sub UserControl_Resize()
'-----------------------------------------------------------------------
'set the various properties of the lblMain here
'-----------------------------------------------------------------------
    On Error Resume Next
    lblMain.Top = (UserControl.ScaleHeight - lblMain.Height) / 2
    lblMain.Left = 0
    lblMain.Width = UserControl.Width - lblMain.Left - 60
End Sub

'-----------------------------------------------------------------------
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'-----------------------------------------------------------------------
'only Caption property is being read by the cotrol
'-----------------------------------------------------------------------
    lblMain.Caption = PropBag.ReadProperty("Caption")
End Sub

'-----------------------------------------------------------------------
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'-----------------------------------------------------------------------
' The control writes the Caption property for this control
'-----------------------------------------------------------------------
    PropBag.WriteProperty "Caption", lblMain.Caption
End Sub

