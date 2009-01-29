VERSION 5.00
Begin VB.UserControl OptionGroup 
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2175
   ScaleHeight     =   495
   ScaleWidth      =   2175
   ToolboxBitmap   =   "OptionGroup.ctx":0000
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   2175
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.Shape shpIn 
         FillStyle       =   0  'Solid
         Height          =   100
         Index           =   0
         Left            =   95
         Shape           =   3  'Circle
         Top             =   170
         Width           =   100
      End
      Begin VB.Shape shpOut 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   200
         Index           =   0
         Left            =   40
         Shape           =   3  'Circle
         Top             =   120
         Width           =   200
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Option Button 1"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   1
         Top             =   120
         Width           =   1890
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "OptionGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   File:       OptionGroup.ctl
'   Author:     Zulfiqar Ahmed, October 2001
'   Purpose:    ActiveX control to display option buttons to be
'               used in Macro 2.2 and above versions.
'-----------------------------------------------------------------------------------
' REVISIONS
'   NCJ 22 Oct 01 - Added OnGroupFocus event, SetFocus method
'               Tidied up some code
'   NCJ 23 Oct 01 - Added ButtonHeight and ButtonWidth properties
' MACRO 3.0
'   NCJ 12 Dec 01 - Changed MouseDown, MouseMove and MouseUp to apply to whole control
'   NCJ 16 Jan 02 - Fixed Enable/Disable as already done in MACRO 2.2
'   NCJ 4 Nov 02 - Added MouseIcon property
'   RS  11/03/2003 - MouseMove/Down/Up events of lblText did take into account the fact
'                    that lblText position was in the containing picMain control, and
'                    returned relative coordinates to picMain and not the UserControl
'                    added picMain position to the raise event
'-----------------------------------------------------------------------------------
Option Explicit

'Constant declarations
Const mnGAP = 60

Private mbManualResize As Boolean
Private mlHighLightColor As Long
Private mlBackColor As Long
Private mlFontColor As Long
Private mbEnabled As Boolean

Dim mnTwipsPerPixel As Integer

'Event Declarations for this control
Public Event Click(OptionIndex As Integer)
Public Event OnGroupFocus()
Public Event DblClick(OptionIndex As Integer)
Public Event KeyDown(OptionIndex As Integer, KeyCode As Integer, Shift As Integer)
Public Event KeyPress(OptionIndex As Integer, KeyAscii As Integer)
Public Event KeyUp(OptionIndex As Integer, KeyCode As Integer, Shift As Integer)
Public Event OnFocus(OptionIndex As Integer)
Public Event ExitFocus(OptionIndex As Integer)
' These events apply to the user control as a whole
' (not the individual option buttons)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

' NB We cannot have DragDrop or DragOver events because
' we are not allowed to pass objects, e.g. the Source


'---------------------------------------------------------------------
Public Sub Init()
'---------------------------------------------------------------------
' Call this procedure before using the control,
' as it sets up some default values for this control
'---------------------------------------------------------------------
    
    mbManualResize = True
    Buttons = 1
    BackColor = Ambient.BackColor
    HighLightColor = vbYellow
    mbEnabled = True
    mlFontColor = lblText(0).ForeColor
    mnTwipsPerPixel = Screen.TwipsPerPixelY

End Sub

'-----------------------------------------------------------------------
Public Property Get ButtonHeight() As Long
'-----------------------------------------------------------------------
' The height of an individual button
'-----------------------------------------------------------------------

    ButtonHeight = picMain(0).Height

End Property

'-----------------------------------------------------------------------
Public Property Let ButtonHeight(lBHt As Long)
'-----------------------------------------------------------------------
' The height of an individual button
' NB this should be set BEFORE setting the Buttons property
' but AFTER setting the Font details (because it needs to calculate the label height)
'-----------------------------------------------------------------------
Dim sglTemp As Single

    picMain(0).Height = lBHt
    lblText(0).Height = TextHeight("W") + 50    ' Fudge factor
    ' Centre the label
    lblText(0).Top = (lBHt - lblText(0).Height) \ 2
    
    ' Centre the circles in the picture box
'    ' NCJ - Nudge each to nearest pixel otherwise it doesn't work properly
    sglTemp = (lBHt - shpOut(0).Height) / 2
    shpOut(0).Top = (sglTemp \ mnTwipsPerPixel) * mnTwipsPerPixel
    
    sglTemp = (shpOut(0).Top + shpOut(0).Height / 2 - shpIn(0).Height / 2)
    shpIn(0).Top = (sglTemp \ mnTwipsPerPixel) * mnTwipsPerPixel
    
End Property

'-----------------------------------------------------------------------
Public Property Get ButtonWidth() As Long
'-----------------------------------------------------------------------
' The Width of an individual button
'-----------------------------------------------------------------------

    ButtonWidth = picMain(0).Width

End Property

'-----------------------------------------------------------------------
Public Property Let ButtonWidth(lBWidth As Long)
'-----------------------------------------------------------------------
' The Width of an individual button
' NB this should be set BEFORE setting the Buttons property
'-----------------------------------------------------------------------
Dim lRadioW As Long

    UserControl.Width = lBWidth
    picMain(0).Width = lBWidth
    
    lRadioW = Me.RadioWidth
    ' Label is what remains to the left of the radio button
    lblText(0).Width = lBWidth - lRadioW
    ' Place the label to the left of the radio
    lblText(0).Left = lRadioW

End Property

'-----------------------------------------------------------------------
Public Property Get RadioWidth() As Long
'-----------------------------------------------------------------------
' Return the width required for the radio button part of the control
' Use this to add to a text width when calculating total button width
'-----------------------------------------------------------------------

    RadioWidth = shpOut(0).Width + shpOut(0).Left + mnGAP

End Property

'-----------------------------------------------------------------------
Public Property Get FontColor() As Long
'-----------------------------------------------------------------------
'Use this property to retrieve the font color for the option buttons
'in the option group control
'-----------------------------------------------------------------------
    
    FontColor = lblText(0).ForeColor

End Property

'-----------------------------------------------------------------------
Public Property Let FontColor(ByVal lForeColor As Long)
'-----------------------------------------------------------------------
'Use this property to set the font color for the option buttons
'in the option group control
'-----------------------------------------------------------------------
Dim nButtons As Integer
    
    For nButtons = 0 To picMain.Count - 1
        lblText(nButtons).ForeColor() = lForeColor
    Next nButtons
    mlFontColor = lForeColor
    PropertyChanged "ForeColor"

End Property

'-----------------------------------------------------------------------
Public Property Get Controls() As Object
'-----------------------------------------------------------------------
'Object enumeration of all controls
'-----------------------------------------------------------------------
    Set Controls = UserControl.Controls
End Property

'-----------------------------------------------------------------------
Public Property Get FontSize() As Byte
'-----------------------------------------------------------------------
'Use this property to retrieve the font size for all option buttons
'in the option group control
'-----------------------------------------------------------------------
    FontSize = lblText(0).FontSize
End Property

'-----------------------------------------------------------------------
Public Property Let FontSize(ByVal byFontSize As Byte)
'-----------------------------------------------------------------------
'Use this property to set the font size for all option buttons
'in the option group control
'-----------------------------------------------------------------------
Dim nButtons As Integer
    
    For nButtons = 0 To lblText.Count - 1
        lblText(nButtons).FontSize = byFontSize
    Next nButtons
    
    PropertyChanged "FontSize"
End Property

'-----------------------------------------------------------------------
Public Property Get FontName() As String
'-----------------------------------------------------------------------
'Use this property to retrieve the font names for all option buttons
'in the option group control
'-----------------------------------------------------------------------
    FontName = lblText(0).FontName
End Property

'-----------------------------------------------------------------------
Public Property Let FontName(ByVal sFontName As String)
'-----------------------------------------------------------------------
Dim nButtons As Integer
    
    For nButtons = 0 To lblText.Count - 1
        lblText(nButtons).Font = sFontName
    Next nButtons
    PropertyChanged "FontName"

End Property

'-----------------------------------------------------------------------
Public Property Set MouseIcon(picMouse As Picture)
'-----------------------------------------------------------------------
' Set the mouse icon for all the labels
' This automatically sets mousepointer to vbCustom as well
'-----------------------------------------------------------------------
Dim nButtons As Integer
    
    ' Set it for all the controls
    For nButtons = 0 To lblText.Count - 1
        picMain(nButtons).MouseIcon = picMouse
        picMain(nButtons).MousePointer = vbCustom
    Next nButtons

End Property

'-----------------------------------------------------------------------
'Use this property to store/retrieve the Italic font property for all option
'buttons in the option group control
'-----------------------------------------------------------------------
Public Property Get FontItalic() As Boolean
    FontItalic = lblText(0).FontItalic
End Property
Public Property Let FontItalic(ByVal bFontItalic As Boolean)
Dim nButtons As Integer
    
    For nButtons = 0 To lblText.Count - 1
        lblText(nButtons).FontItalic = bFontItalic
    Next nButtons
End Property

'-----------------------------------------------------------------------
'Use this property to store/retrieve the font bold property for all option
'buttons in the option group control
'-----------------------------------------------------------------------
Public Property Get FontBold() As Boolean
    FontBold = lblText(0).FontBold
End Property
Public Property Let FontBold(ByVal bFontBold As Boolean)
Dim nButtons As Integer
    
    For nButtons = 0 To lblText.Count - 1
        lblText(nButtons).FontBold = bFontBold
    Next nButtons
End Property
'-----------------------------------------------------------------------
'Use this property to store/retrieve the highlight color for the option
'buttons in the option group control
'-----------------------------------------------------------------------
Public Property Get HighLightColor() As Long
    HighLightColor = mlHighLightColor
End Property
Public Property Let HighLightColor(ByVal lHighLightColor As Long)
    mlHighLightColor = lHighLightColor
End Property

'-----------------------------------------------------------------------
'Use this property to store/retrieve the back color for the option
'buttons in the option group control
'-----------------------------------------------------------------------
Public Property Get BackColor() As Long
    BackColor = mlBackColor
End Property
Public Property Let BackColor(ByVal lBackColor As Long)
    mlBackColor = lBackColor
    picMain(0).BackColor = lBackColor
    PropertyChanged "BackColor"
End Property

'-----------------------------------------------------------------------
Public Property Get SelectedItem() As Integer
'-----------------------------------------------------------------------
'Use this property to retrieve the currently selected option button
'in the option group control
' Returns -1 if no option is selected
'-----------------------------------------------------------------------
Dim nCount As Integer

    SelectedItem = -1
    For nCount = 0 To Buttons
        If shpIn(nCount).Visible = True Then
            SelectedItem = nCount
            Exit For
        End If
    Next nCount

End Property

'-----------------------------------------------------------------------
Public Property Get Buttons() As Byte
'-----------------------------------------------------------------------
'Use this property to retrieve the max. index for option buttons
'in the option group control
' This is Button Count - 1
'-----------------------------------------------------------------------
    
    Buttons = picMain.Count - 1

End Property

'-----------------------------------------------------------------------
Public Property Let Buttons(ByVal byButtons As Byte)
'-----------------------------------------------------------------------
'Use this property to set the number of  option buttons
'in the option group control
'-----------------------------------------------------------------------
Dim nOptions As Integer
Dim nPrevCount As Integer

    nPrevCount = picMain.Count
    mbManualResize = False
    If byButtons = 0 Then
        Err.Raise 6500, "IMEDOptionControl", "Option indexes must be greater than 0"
    Else
        UserControl.Height = byButtons * picMain(0).Height
        If byButtons > nPrevCount Then
            For nOptions = nPrevCount To byButtons - 1
                'load 1 pic, 2 shapes and 1 lable
                Load picMain(nOptions)
                Load shpOut(nOptions)
                Load shpIn(nOptions)
                Load lblText(nOptions)
                
                'set left, top, width, height and visibility properties for picture box
                picMain(nOptions).Left = picMain(nOptions - 1).Left
                picMain(nOptions).Top = picMain(nOptions - 1).Top + picMain(nOptions - 1).Height
                picMain(nOptions).Width = UserControl.Width
                picMain(nOptions).Height = picMain(0).Height
                picMain(nOptions).Visible = True
                
                'set left, top, width, height and visibility properties for label control
                Set lblText(nOptions).Container = picMain(nOptions)
                lblText(nOptions).Left = lblText(0).Left
                lblText(nOptions).Top = lblText(0).Top
                lblText(nOptions).Height = lblText(0).Height
                lblText(nOptions).Width = lblText(0).Width
                lblText(nOptions).Caption = "Option Button " & nOptions + 1
                lblText(nOptions).Visible = True
                
                
                'set left, top, width, height and visibility properties for outer shape
                Set shpOut(nOptions).Container = picMain(nOptions)
                shpOut(nOptions).Left = shpOut(0).Left
                shpOut(nOptions).Top = shpOut(0).Top
                shpOut(nOptions).Height = shpOut(0).Height
                shpOut(nOptions).Width = shpOut(0).Width
                shpOut(nOptions).Visible = True
                
                'set left, top, width, height and visibility properties for inner shape
                Set shpIn(nOptions).Container = picMain(nOptions)
                shpIn(nOptions).Left = shpIn(0).Left
                shpIn(nOptions).Top = shpIn(0).Top
                shpIn(nOptions).Height = shpIn(0).Height
                shpIn(nOptions).Width = shpIn(0).Width
                shpIn(nOptions).Visible = True
                
                picMain(nOptions).TabIndex = nOptions
                picMain(nOptions).TabStop = False
                picMain(nOptions).BackColor = BackColor
                

            Next nOptions
        End If
    End If
    PropertyChanged "Buttons"
End Property

'-----------------------------------------------------------------------
Public Sub SetFocus()
'-----------------------------------------------------------------------
' Set the focus to the selected option button
' or to the first if none is selected
'-----------------------------------------------------------------------
Dim nIndex As Integer

    nIndex = Me.SelectedItem
    If nIndex > -1 Then
        picMain(nIndex).SetFocus
    Else
        ' Focus to first one
        picMain(0).SetFocus
    End If

End Sub

'-----------------------------------------------------------------------
Private Sub lblText_Click(Index As Integer)
'-----------------------------------------------------------------------
' Clicking the label is the same as clicking the button
'-----------------------------------------------------------------------
    
    Call picMain(Index).SetFocus
    Call picMain_Click(Index)

End Sub


'-----------------------------------------------------------------------
Private Sub picMain_Click(Index As Integer)
'-----------------------------------------------------------------------
' Select the option button when they click
' Only pass on the click event if the value changed
'-----------------------------------------------------------------------
    
    If SetOptionValue(Index) Then
        Me.Refresh
        RaiseEvent Click(Index)
    End If

End Sub


'-----------------------------------------------------------------------
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'-----------------------------------------------------------------------
    
    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

'-----------------------------------------------------------------------
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'-----------------------------------------------------------------------

    Debug.Print "UserControl_MouseMove", X, Y, "Compensated: " & X, Y
    RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub lblText_DblClick(Index As Integer)
    RaiseEvent DblClick(Index)
End Sub

'-----------------------------------------------------------------------
Private Sub lblText_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'-----------------------------------------------------------------------
' Offset X and Y to represent User Control coordinates
'-----------------------------------------------------------------------
    
    ' RS 11/03/2003: lblText position is in its container, picMain(Index), add position of picMain(Index) to get position in entire usercontrol
    RaiseEvent MouseDown(Button, Shift, X + lblText(Index).Left + picMain(Index).Left, Y + lblText(Index).Top + picMain(Index).Top)

End Sub

'-----------------------------------------------------------------------
Private Sub lblText_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'-----------------------------------------------------------------------
' Offset X and Y to represent User Control coordinates
'-----------------------------------------------------------------------
    
    
    ' RS 11/03/2003: The position of lblText(index) is very different from the corresponding picMain(index) element
    ' Debug.Print "lblText_MouseMove", Index, X, Y, "Position:", lblText(Index).Left, lblText(Index).Top, "Compensated:", X + lblText(Index).Left, Y + lblText(Index).Top
    RaiseEvent MouseMove(Button, Shift, X + lblText(Index).Left + picMain(Index).Left, Y + lblText(Index).Top + picMain(Index).Top)

End Sub

'---------------------------------------------------------------------
Private Sub lblText_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
' Offset X and Y to represent User Control coordinates
'-----------------------------------------------------------------------

    ' RS 11/03/2003: lblText position is in its container, picMain(Index), add position of picMain(Index) to get position in entire usercontrol
    RaiseEvent MouseUp(Button, Shift, X + lblText(Index).Left + picMain(Index).Left, Y + lblText(Index).Top + picMain(Index).Top)
    
End Sub


'---------------------------------------------------------------------
Private Sub picMain_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
' Offset X and Y to represent User Control coordinates
'-----------------------------------------------------------------------
    
    RaiseEvent MouseUp(Button, Shift, X + picMain(Index).Left, Y + picMain(Index).Top)

End Sub

Private Sub picMain_DblClick(Index As Integer)
    RaiseEvent DblClick(Index)
End Sub

Private Sub picMain_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(Index, KeyCode, Shift)
End Sub

'-----------------------------------------------------------------------
Private Sub picMain_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'-----------------------------------------------------------------------
' Offset X and Y to represent User Control coordinates
'-----------------------------------------------------------------------
    
    RaiseEvent MouseDown(Button, Shift, X + picMain(Index).Left, Y + picMain(Index).Top)

End Sub

'-----------------------------------------------------------------------
Private Sub picMain_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'-----------------------------------------------------------------------
' Offset X and Y to represent User Control coordinates
'-----------------------------------------------------------------------
    
    RaiseEvent MouseMove(Button, Shift, X + picMain(Index).Left, Y + picMain(Index).Top)

End Sub
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------

'-----------------------------------------------------------------------
Private Sub picMain_GotFocus(Index As Integer)
'-----------------------------------------------------------------------
' Raise OnFocus event here.
' Also here we set the HighLightColor property defined by the user.
'-----------------------------------------------------------------------
    
    picMain(Index).BackColor = HighLightColor
    RaiseEvent OnFocus(Index)

End Sub

'-----------------------------------------------------------------------
Private Sub picMain_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'-----------------------------------------------------------------------
' Intercept arrow up and arrow down keys
' otherwise pass the event on
'-----------------------------------------------------------------------
    
    If KeyCode = vbKeyDown Then
        SetControlDown (Index)
    ElseIf KeyCode = vbKeyUp Then
        SetControlUp (Index)
    Else
        RaiseEvent KeyDown(Index, KeyCode, Shift)
    End If
    
End Sub

'-----------------------------------------------------------------------
Private Sub picMain_KeyPress(Index As Integer, KeyAscii As Integer)
'-----------------------------------------------------------------------
' Treat Space key as same as Click
' otherwise pass event on
'-----------------------------------------------------------------------
    
    If KeyAscii = vbKeySpace Then
        ' Same as click
        Call picMain_Click(Index)
    Else
        RaiseEvent KeyPress(Index, KeyAscii)
    End If

End Sub

'-----------------------------------------------------------------------
Private Sub picMain_LostFocus(Index As Integer)
'-----------------------------------------------------------------------
'When the focus from an picture box is lost, we replace the HighLightColor
'with the BackColor.
'-----------------------------------------------------------------------
    
    picMain(Index).BackColor = BackColor
    RaiseEvent ExitFocus(Index)

End Sub

'---------------------------------------------------------------------------
Private Sub SetControlDown(iIndex As Integer)
'---------------------------------------------------------------------------
'This procedure is called when user is pressing the down arrow key
'This procedure sets the appropriate focus in response to that key
'---------------------------------------------------------------------------
    
    If iIndex = picMain.Count - 1 Then
        ' Go from last to first
        picMain(0).SetFocus
    Else
        ' Go to next one
        picMain(iIndex + 1).SetFocus
    End If
End Sub

'---------------------------------------------------------------------------
Private Sub SetControlUp(iIndex As Integer)
'---------------------------------------------------------------------------
'This procedure is called when user is pressing the up arrow key
'This procedure sets the appropriate focus in response to that key
'---------------------------------------------------------------------------
    
    If iIndex = 0 Then
        ' Go from first to last
        picMain(picMain.Count - 1).SetFocus
    Else
        ' Go to previous one
        picMain(iIndex - 1).SetFocus
    End If

End Sub

'---------------------------------------------------------------------
Public Sub UnselectAll()
'---------------------------------------------------------------------
' Clear the selected option button from the group
'---------------------------------------------------------------------
Dim nCount As Integer
    
    For nCount = 0 To picMain.Count - 1
        shpIn(nCount).Visible = False
       ' shpOut(nCount).FillStyle = vbSolid
       ' shpOut(nCount).FillColor = vbWhite
    Next nCount
    picMain(0).BackColor = mlBackColor

End Sub

'---------------------------------------------------------------------
Private Function SetOptionValue(nIndex As Integer) As Boolean
'---------------------------------------------------------------------
' Select the given option button
' Return TRUE if it changed, or FALSE if it didn't change
' Hide all the inner shapes and then make the selected one visible
'---------------------------------------------------------------------
Dim nCount As Integer
        
    SetOptionValue = False
    If shpIn(nIndex).Visible Then
        ' Already selected
        Exit Function
    Else
        For nCount = 0 To picMain.Count - 1
            shpIn(nCount).Visible = False
        Next nCount
        
        shpIn(nIndex).Visible = True
        shpIn(nIndex).ZOrder 0
        SetOptionValue = True
    End If

End Function

'---------------------------------------------------------------------
Public Property Let Caption(Index As Integer, sCaption As String)
'---------------------------------------------------------------------
'set the appropriate captions for an option control. Raise an error
'if the control element does not exist
'---------------------------------------------------------------------
    
    Call IndexError(Index)
    lblText(Index).Caption = sCaption
    
End Property

'---------------------------------------------------------------------
Public Property Get Caption(Index As Integer) As String
'---------------------------------------------------------------------
'retrieves the appropriate captions for an option control. Raise an error
'if the control element does not exist
'---------------------------------------------------------------------
    
    Call IndexError(Index)
    Caption = lblText(Index).Caption

End Property

'-----------------------------------------------------------------------
Public Property Let Enabled(ByVal bEnabled As Boolean)
'-----------------------------------------------------------------------
'Use this property to set the enabled property of option group control
'-----------------------------------------------------------------------
Dim nCount As Integer
Dim nButtons As Integer
Const lGREY = &HC0C0C0

    nButtons = picMain.Count - 1
    
    For nCount = 0 To nButtons
        If bEnabled Then
            picMain(nCount).Enabled = True
            lblText(nCount).Enabled = True
            shpOut(nCount).FillColor = vbWhite
            shpOut(nCount).FillStyle = vbSolid
        Else
            picMain(nCount).Enabled = False
            lblText(nCount).Enabled = False
            shpOut(nCount).FillColor = lGREY
            shpOut(nCount).FillStyle = vbSolid
        End If

   Next nCount
   
   ' NCJ 16 Jan 02 - Ensure entire control is enabled/disabled
   UserControl.Enabled = bEnabled
    
   mbEnabled = bEnabled

End Property

'-----------------------------------------------------------------------
Public Property Get Enabled() As Boolean
'-----------------------------------------------------------------------
'Use this property to retrieve the enabled property of option group control
'-----------------------------------------------------------------------
    
    Enabled = mbEnabled

End Property

'---------------------------------------------------------------------
Public Property Let TagValue(Index As Integer, vTag As Variant)
'---------------------------------------------------------------------
'sets the tag for an option button and raise an error if the index
'doesn't exist
'---------------------------------------------------------------------
    
    Call IndexError(Index)
    lblText(Index).Tag = vTag

End Property

'---------------------------------------------------------------------
Public Property Get TagValue(Index As Integer) As Variant
'---------------------------------------------------------------------
'retrieves the tag for an option button and raise an error if the index
'doesn't exist
'---------------------------------------------------------------------
    
    Call IndexError(Index)
    TagValue = lblText(Index).Tag

End Property

'---------------------------------------------------------------------
Private Sub UserControl_GotFocus()
'---------------------------------------------------------------------
' Pass through the GotFocus event for the whole group
' NCJ 23 Oct 01 - THIS EVENT DOESN'T ACTUALLY HAPPEN!
'---------------------------------------------------------------------

    RaiseEvent OnGroupFocus
    
End Sub


'---------------------------------------------------------------------
Private Sub UserControl_Resize()
'---------------------------------------------------------------------
'Resize the option buttons as the user control is being resized either
'at design or at run time
'---------------------------------------------------------------------
Dim nCounter As Integer

    If mbManualResize = False Then Exit Sub
    If picMain.Count = 1 Then
        picMain(0).Move 0, 0, UserControl.Width, UserControl.Height
        
    Else
        For nCounter = 1 To picMain.Count - 1
            picMain(nCounter).Move 0, 0, UserControl.Width, UserControl.Height / nCounter
        Next nCounter
    End If
    
End Sub

'---------------------------------------------------------------------
Public Function About() As Variant
Attribute About.VB_UserMemId = -552
'---------------------------------------------------------------------
'Display About dialogue for this control
'---------------------------------------------------------------------
    Load frmAbout
    frmAbout.Show vbModal
End Function

'---------------------------------------------------------------------
Public Function IsSelected(Index As Integer) As Boolean
'---------------------------------------------------------------------
'returns if the the option button is selected or not and raise an error if the index
'doesn't exist
'---------------------------------------------------------------------
    
    Call IndexError(Index)
    IsSelected = shpIn(Index).Visible

End Function

'---------------------------------------------------------------------
Public Sub SetSelected(Index As Integer)
'---------------------------------------------------------------------
'set the the option button value to selected and raise an error if the index
'doesn't exist
'---------------------------------------------------------------------
    
    Call IndexError(Index)
    SetOptionValue (Index)

End Sub

'-----------------------------------------------------------------------
Public Sub Refresh()
'-----------------------------------------------------------------------
'Use this method if the control is failing to repaint itself
'-----------------------------------------------------------------------
    
    UserControl.Refresh
    
End Sub

'---------------------------------------------------------------------
Public Property Let ToolTipValue(Index As Integer, sToolTip As String)
'---------------------------------------------------------------------
'sets the tooltiptext property for an option button and raise an error if
'the index doesn't exist
'---------------------------------------------------------------------
    
    Call IndexError(Index)
    ' Set for both picture and label
    lblText(Index).ToolTipText = sToolTip
    picMain(Index).ToolTipText = sToolTip

End Property

'---------------------------------------------------------------------
Public Property Get ToolTipValue(Index As Integer) As String
'---------------------------------------------------------------------
'sets the tooltiptext property for an option button and raise an error if
'the index doesn't exist
'---------------------------------------------------------------------------
    
    Call IndexError(Index)
    ToolTipValue = lblText(Index).ToolTipText

End Property

'---------------------------------------------------------------------------
Private Sub IndexError(nIndex As Integer)
'---------------------------------------------------------------------------
' Raise an error if nIndex is greater than the Button count
'---------------------------------------------------------------------------

    If nIndex > Buttons Then
        Err.Raise 6550, "IMEDOptionControl", "Options index " & nIndex & " does not exist"
    End If

End Sub
