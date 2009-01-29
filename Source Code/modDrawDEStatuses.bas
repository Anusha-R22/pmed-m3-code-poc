Attribute VB_Name = "modDrawDEStatuses"
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   File:       modDrawDEStatuses.bas
'   Author:     Zulfi Ahmed/Toby Aldridge, August 2002
'   Purpose:    Handle drawing of Statuses in Data Entry eForms
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   ZA 28/08/2002: Non MACRO specific drawing code
'   TA 28/08/2002: MACRO specific function wrappers
'   NCJ 30 Aug 02 - Debugging
' MLM 16/09/02: Allow controls without 3D border.
'----------------------------------------------------------------------------------------'

Option Explicit

Private Const m_SDV_GREEN = 45056  '49152 is lighter, '32768 is darker

Private Enum eBorderStyle
    BSTransparent = 0
    BSSolid = 1
    BSDash = 2
    BSDot = 3
    BSDashDot = 4
    BSDashDotDot = 5
    BSInsideSolid = 6
End Enum

'----------------------------------------------------------------------------------------'
Public Sub SetChangeCountGraphics(oControl As Control, _
                                oImg As Image, nChangeCount As Integer)
'----------------------------------------------------------------------------------------'
' Draw graphic according to ChangeCount
' nChangeCount is how many (non-"Missing") response rows are in the DIRH table
'----------------------------------------------------------------------------------------'
Dim nPrevResponses As Integer

    On Error GoTo ErrLabel

    ' Don't include current response in "Prev. responses"
    ' (so subtract 1 from ChangeCount if greater than 0)
    nPrevResponses = 0
    If nChangeCount > 1 Then
        nPrevResponses = nChangeCount - 1
    End If
    
    Select Case nPrevResponses
        Case 0
            oImg.Visible = False
        'set the small bar
        Case 1 'previous value
            oImg.Visible = True
        'set the small bar
        Case 2 'previous values
            oImg.Visible = True
        Case Is > 2 'several previous values
            oImg.Visible = True
    End Select
  
Exit Sub

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modDrawDEStatuses.SetChangeCountGraphics"

End Sub

'----------------------------------------------------------------------------------------'
Public Sub SetSDVStatusGraphics(oControl As Control, oSHPControl As Shape, _
                                enSDVStatus As MACRODEBS30.eSDVStatus, _
                                nDefaultBorderStyle As Integer)
'----------------------------------------------------------------------------------------'
'draw graphic according to SDV status
'
'MLM 16/09/02: Added nDefaultBorderStyle parameter. If there are no SDVs, apply this style.
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrLabel

    Select Case enSDVStatus
    Case MACRODEBS30.eSDVStatus.ssCancelled, MACRODEBS30.eSDVStatus.ssNone
        If TypeOf oControl Is TextBox Then
            'turn border back on if text box
            oControl.BorderStyle = nDefaultBorderStyle
        End If
        oSHPControl.Visible = False
    Case MACRODEBS30.eSDVStatus.ssPlanned
        Call DrawBorder(BSDot, m_SDV_GREEN, oControl, oSHPControl)
    Case MACRODEBS30.eSDVStatus.ssQueried
        Call DrawBorder(BSDash, vbRed, oControl, oSHPControl)
    Case MACRODEBS30.eSDVStatus.ssComplete
        Call DrawBorder(BSSolid, m_SDV_GREEN, oControl, oSHPControl)
    End Select
    
Exit Sub

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modDrawDEStatuses.SetSDVStatusGraphics"

End Sub


'----------------------------------------------------------------------------------------'
Private Sub DrawBorder(ByVal enStyle As eBorderStyle, ByVal lColour As Long, _
                        ByRef oControl As Control, ByRef oSHPControl As Shape)
'----------------------------------------------------------------------------------------'
' draws custom border with customer colour
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrLabel
    
    If TypeOf oControl Is TextBox Then
        'if a textbox then turn off border
        oControl.BorderStyle = 0 'no border
    End If
    
    With oSHPControl
        If TypeOf oControl Is MSFlexGrid Then
            'when called from schedule
            .Left = oControl.Left + oControl.CellLeft - 10
            .Top = oControl.Top + oControl.CellTop - 10
            .Width = oControl.CellWidth + 30
            .Height = oControl.CellHeight + 30
        Else
            .Left = oControl.Left - 10
            .Top = oControl.Top - 10
            .Width = oControl.Width + 30
            .Height = oControl.Height + 30
        End If
        .BorderStyle = enStyle
        If enStyle = BSSolid Then
            .BorderWidth = 2
        Else
            .BorderWidth = 1
        End If
        .BorderColor = lColour '154689
        .Visible = True
   End With
    
Exit Sub

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modDrawDEStatuses.DrawBorder"
  
End Sub

