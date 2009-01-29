VERSION 5.00
Begin VB.UserControl DownArrowButton 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   510
   ScaleHeight     =   420
   ScaleWidth      =   510
   ToolboxBitmap   =   "DownArrowButton.ctx":0000
   Begin VB.Image imgMain 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   0
      Picture         =   "DownArrowButton.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   225
   End
End
Attribute VB_Name = "DownArrowButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   File:       DownArrowButton.ctl
'   Author:     Zulfiqar Ahmed, October 2001
'   Purpose:    ActiveX control to display a button with down arrow image, that is
'               to be used in Macro 2.2 and above versions.
'
'-----------------------------------------------------------------------------------

Option Explicit

'declare the following events here
Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

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
Private Sub imgMain_Click()
'-----------------------------------------------------------------------
'Raise Click event here
'-----------------------------------------------------------------------
    RaiseEvent Click
End Sub

'-----------------------------------------------------------------------
Private Sub imgMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'-----------------------------------------------------------------------
'Move the imgMain so that user can see the click on this control. Also raise
'MouseDown event here
'-----------------------------------------------------------------------
    imgMain.Left = 25
    imgMain.Top = 25
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

'-----------------------------------------------------------------------
Private Sub imgMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'-----------------------------------------------------------------------
' Raise Mouse Move event here
'-----------------------------------------------------------------------
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub


'-----------------------------------------------------------------------
Private Sub imgMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'-----------------------------------------------------------------------
'Move the imgMain to its position when the user releases the mouse. Also
'raise MouseUp event here
'-----------------------------------------------------------------------
    imgMain.Left = 0
    imgMain.Top = 0
    RaiseEvent MouseMove(Button, Shift, X, Y)
    
End Sub

Public Sub Init()
    BackColor = Ambient.BackColor
End Sub

'-----------------------------------------------------------------------
Private Sub UserControl_Resize()
'-----------------------------------------------------------------------
'Don't let the user resize the control during desing or run time. We want
' a fixed width and height of this control
'-----------------------------------------------------------------------
    UserControl.Width = 300
    UserControl.Height = 300
End Sub

'-----------------------------------------------------------------------
Public Property Let BackColour(Colour As OLE_COLOR)
'-----------------------------------------------------------------------
'Use this property to store the back color for the Down Arrow
'buttons in the IMED controls
'-----------------------------------------------------------------------
    UserControl.BackColor = Colour
End Property

'-----------------------------------------------------------------------
Public Property Get BackColour() As OLE_COLOR
'-----------------------------------------------------------------------
'Use this property to retrieve the back color for the Down Arrow
'buttons in IMED controls
'-----------------------------------------------------------------------
    BackColour = UserControl.BackColor
End Property

'-----------------------------------------------------------------------
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'-----------------------------------------------------------------------
'only BackColour property is beign read by the cotrol
'-----------------------------------------------------------------------
    UserControl.BackColor = PropBag.ReadProperty("BackColour", UserControl.BackColor)
End Sub

'-----------------------------------------------------------------------
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'-----------------------------------------------------------------------
' The control writes the BackColour property for this control
'-----------------------------------------------------------------------
    PropBag.WriteProperty "BackColour", UserControl.BackColor
End Sub

