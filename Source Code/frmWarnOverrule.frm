VERSION 5.00
Begin VB.Form frmWarnOverrule 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Overrule"
   ClientHeight    =   5700
   ClientLeft      =   8145
   ClientTop       =   5655
   ClientWidth     =   6345
   ControlBox      =   0   'False
   Icon            =   "frmWarnOverrule.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   6345
   Begin VB.TextBox txtReason 
      Height          =   795
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmWarnOverrule.frx":000C
      Top             =   3480
      Width           =   6195
   End
   Begin VB.ListBox lstReason 
      Height          =   1230
      Left            =   60
      TabIndex        =   5
      Top             =   4395
      Width           =   6195
   End
   Begin VB.Label lblResponse 
      Caption         =   "Value: response value here"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   330
      TabIndex        =   8
      Top             =   300
      Width           =   5835
   End
   Begin VB.Label lblReason 
      Caption         =   "Overrule reason"
      Height          =   195
      Left            =   60
      TabIndex        =   7
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lcmdOverrule 
      AutoSize        =   -1  'True
      Caption         =   "Overrule this warning"
      Height          =   195
      Left            =   960
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2940
      Width           =   1485
   End
   Begin VB.Label lblErrorMessage 
      AutoSize        =   -1  'True
      Caption         =   "lblErrorMessage"
      Height          =   1695
      Left            =   720
      TabIndex        =   3
      Top             =   840
      Width           =   5505
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblErrorType 
      Caption         =   "lblErrorType"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   600
      Width           =   3360
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Left            =   60
      Top             =   840
      Width           =   240
   End
   Begin VB.Label lcmdClose 
      Alignment       =   1  'Right Justify
      Caption         =   "Save and close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4890
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   60
      Width           =   1335
   End
   Begin VB.Label lblQuestion 
      Caption         =   "Question: question name here"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4755
   End
End
Attribute VB_Name = "frmWarnOverrule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000-2001. All Rights Reserved
'   File:       frmWarnOverrule.frm
'   Author:     Toby Aldridge Sept 2001
'   Purpose:    Allow user to overrule / unoverrule warnings
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   TA 22/01/2001 SR4128: disallow input of invalid characters ("`|~)
'   ZA 14/08/01, Add new parameters into the Display function
'   NCJ 30/8/01 - Corrected Display parameters
'   DPH - 12/10/2001 New case to allow for reject messages
'   ASH - 19/03/2002 Modified routines View and Display
'   ZA  - 15/05/2002 Modified form to add validate button + icon
'   NCJ 15 Aug 02 - Changed Display parameters and added Repeat Number to dialog caption
'   NCJ 26 Sept 02 - Do not allow warning overrules/edits on read-only eForms
'   NCJ 30 Oct 02 - Include question name in error message
'   TA 01/11/2002 - Changed to new icons
'   TA 04/12/2002 - Start to change UI
'   NCJ 6 Jan 03 - Debugged code and made it work
'   TA 21/01/03: ok/cancel/enter/close all save changes - there is no cancel
'   NCJ 28 Mar 03 - Deal with ampersands in error message (double them up)
'   NCJ 2 Sept 03 - User must have ChangeData rights to overrule or edit overrule
'   NCJ 15 Jul 04 - Bug 2291 - Ensure ChangeData rights work correctly!
'   TA  08/08/2005: setting of Font.Charset = 1 to allow eastern european characters. CBD2591.
'----------------------------------------------------------------------------------------'

Option Explicit

Private mnStatus As eStatus
Private msOverruleReason As String
Private mbChanged As Boolean
Private msValidationMessage As String
Private moResponse As Response

' NCJ 6 Jan 03 - True while loading form
Private mbLoading As Boolean

'---------------------------------------------------------------------
Public Function Display(oResponse As Response, _
                    ByVal nStatus As Integer, _
                    ByVal sResponseValue As String, _
                    ByVal sValidationMessage As String, _
                    ByRef sReason As String, oCoord As clsCoords) As Long
'---------------------------------------------------------------------
'   Display form
'   Output:
'       sReason - string to receive overrule reason
'       function - status (-1 for no change/cancel)
'---------------------------------------------------------------------
' NCJ 15 Aug 02 - Changed parameters
'---------------------------------------------------------------------
Dim colRFOs As Collection
Dim i As Long
   
    mbLoading = True
    
    Me.Top = oCoord.Row
    Me.Left = oCoord.Col
    
    Set moResponse = oResponse
    mnStatus = nStatus
    msOverruleReason = sReason
    msValidationMessage = sValidationMessage
    
    ' Get Reasons For Overrule from StudyDef
    Set colRFOs = oResponse.EFormInstance.eForm.Study.RFOs
    lstReason.Enabled = (colRFOs.Count > 0)
    lstReason.Clear
    For i = 1 To colRFOs.Count
        lstReason.AddItem colRFOs(i)
    Next
    
    lblQuestion.Caption = "Question: " & QuestionNameText(oResponse)
    lblResponse.Caption = "Value: " & sResponseValue
    
    Me.Caption = "    "
    mbChanged = False
    
    
    Call View
    
    mbLoading = False
    
    'ensure we appear completely on the screen
    If Me.Top < 100 Then
        Me.Top = 100
    End If
    If Me.Top + (lstReason.Top + Me.Height - Me.ScaleHeight) + lstReason.Height + 120 > Screen.Height - 500 Then
        Me.Top = Screen.Height - (lstReason.Top + Me.Height - Me.ScaleHeight) - (lstReason.Height + 120) - 500
    End If
    If Me.Left < 100 Then
        Me.Left = 100
    End If
    If Me.Left + Me.Width > Screen.Width Then
        Me.Left = Screen.Width - Me.Width
    End If
    
    'TA 2/4/2003: turn off hourglass while showing modal form
    HourglassSuspend
    Me.Show vbModal
    HourglassResume
    
    If mbChanged Then
        ' Return what they typed
        sReason = msOverruleReason
        Display = mnStatus
    Else
        Display = -1
    End If

End Function


'---------------------------------------------------------------------
Private Sub lcmdMore_Click()
'---------------------------------------------------------------------
'Allow user to view validation message window
'---------------------------------------------------------------------

    frmValidation.Display moResponse.Element
    
End Sub


'---------------------------------------------------------------------
Private Sub lcmdclose_Click()
'---------------------------------------------------------------------

    Unload Me
    
End Sub

'---------------------------------------------------------------------
Private Sub lcmdOverrule_Click()
'---------------------------------------------------------------------
'   overrule / remove overrule
'---------------------------------------------------------------------
   
    If mnStatus = eStatus.OKWarning Then
        If DialogQuestion("Undo overrule on question " & vbCrLf & vbCrLf & "     " & QuestionNameText(moResponse) & vbCrLf & vbCrLf & "and reset status to Warning?", "Undo Overrule") = vbYes Then
            msOverruleReason = ""
            mnStatus = eStatus.Warning
            mbChanged = True
        End If
    Else
        Me.Height = (Me.Height - Me.ScaleHeight) + lstReason.Top + lstReason.Height + 120
        'TA 22/01/2001 SR4128: disallow input of invalid characters ("`|~)
  '     If frmInputBox.Display("Overrule", "Reason for overrule", msOverruleReason, , , , valOnlySingleQuotes) Then
  '     End If
        mnStatus = eStatus.OKWarning
        mbChanged = True
    End If
    
    Call View
    
End Sub

'---------------------------------------------------------------------
Private Sub View()
'---------------------------------------------------------------------
'  Set up labels, buttons etc. according to status of question
' NCJ 2 Sept 03 - User must have ChangeData rights
'---------------------------------------------------------------------
Dim bCanOverrule As Boolean

    'ash 19/03/2002
    ' NCJ 26 Sept 02 - Must take into account read-onlyness of EFI
    ' NCJ 2 Sept 03 - User must have Change Data permission too
    bCanOverrule = (goUser.CheckPermission(gsFnOverruleWarnings) _
                    And goUser.CheckPermission(gsFnChangeData) _
                    And (moResponse.LockStatus = eLockStatus.lsUnlocked) _
                    And (Not moResponse.EFormInstance.ReadOnly))
    
    ' NCJ 30 Oct 02 - New routine to generate error message text
    lblErrorMessage.Caption = GetErrorMessageText(moResponse, mnStatus, msValidationMessage)
    
    ' NCJ 15 Jul 04 - These fields default to disabled
    txtReason.Enabled = False
    lstReason.Enabled = False
    
    Select Case mnStatus
    Case eStatus.Inform
        lblErrorType.Caption = " Data Inform "
        lcmdOverrule.Visible = False
        txtReason.Text = ""
        imgIcon.Picture = frmImages.imglistStatus.ListImages(DM30_ICON_INFORM).Picture
        lcmdClose.Caption = "Close"
    Case eStatus.Warning
        lblErrorType.Caption = " Warning "
        lcmdOverrule.Caption = "Overrule this warning"
        lcmdOverrule.Visible = bCanOverrule
'        txtReason.Enabled = bCanOverrule
'        lstReason.Enabled = bCanOverrule
        txtReason.Text = ""
        imgIcon.Picture = frmImages.imglistStatus.ListImages(DM30_ICON_WARNING).Picture
        lcmdClose.Caption = "Save and close"
     Case eStatus.OKWarning
        lblErrorType.Caption = " Warning "
        lcmdOverrule.Caption = "Remove overrule reason"
        lcmdOverrule.Visible = bCanOverrule
        txtReason.Enabled = bCanOverrule
        lstReason.Enabled = bCanOverrule
        txtReason.Text = msOverruleReason
        If msOverruleReason = "" Then
            imgIcon.Picture = frmImages.imglistStatus.ListImages(DM30_ICON_WARNING).Picture
        Else
            imgIcon.Picture = frmImages.imglistStatus.ListImages(DM30_ICON_OK_WARNING).Picture
        End If
        lcmdClose.Caption = "Save and close"
        ' DPH - 12/10/2001 New case for reject messages
     Case eStatus.InvalidData
        lblErrorType.Caption = " Data Invalid "
        lcmdOverrule.Visible = False
        txtReason.Text = ""
        imgIcon.Picture = frmImages.imglistStatus.ListImages(DM30_ICON_INVALID).Picture
        lcmdClose.Caption = "Close"
    End Select


    'TA 10/04/2003: disable reasons for overrule when warning
    ' NCJ 15 Jul 04 - Enabling according to bCanOverrule already done (Bug 2291)
    If mnStatus = eStatus.OKWarning Or mnStatus = eStatus.Warning Then
        lcmdOverrule.Top = lblErrorMessage.Top + lblErrorMessage.Height + 120
        lblReason.Top = lcmdOverrule.Top + lcmdOverrule.Height + 120
        txtReason.Top = lblReason.Top + lblReason.Height + 30
        lstReason.Top = txtReason.Top + txtReason.Height + 120
        If mnStatus = eStatus.Warning Then
            Me.Height = (Me.Height - Me.ScaleHeight) + lcmdOverrule.Top + lcmdOverrule.Height + 120
'            txtReason.Enabled = False
'            lstReason.Enabled = False
        Else
            Me.Height = (Me.Height - Me.ScaleHeight) + lstReason.Top + lstReason.Height + 120
'            txtReason.Enabled = True
'            lstReason.Enabled = True
        End If
            
    Else
        Me.Height = (Me.Height - Me.ScaleHeight) + lblErrorMessage.Top + lblErrorMessage.Height + 120
    End If


End Sub

'---------------------------------------------------------------------
Private Function GetErrorMessageText(oResponse As Response, enStatus As eStatus, sErrMsg As String) As String
'---------------------------------------------------------------------
' Get the error message text to display in the window
'---------------------------------------------------------------------
Dim sText As String

    Select Case enStatus
    Case eStatus.Warning, eStatus.OKWarning
        sText = "The following warning for question " & QuestionNameText(moResponse) & " has been generated:"
    Case eStatus.Inform
        sText = "The following flag for question " & QuestionNameText(oResponse) & " has been generated:"
    Case eStatus.InvalidData
        sText = "The value for question " & QuestionNameText(oResponse) & " has been rejected because:"
    End Select
    
    ' NCJ 28 Mar 03 - Prevent ampersands from being interpreted by VB
    GetErrorMessageText = sText & vbCrLf & vbCrLf & Replace(sErrMsg, "&", "&&")
    
End Function

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------

'---------------------------------------------------------------------
Dim oControl As Control

    Me.BackColor = eMACROColour.emcBackground

        
    On Error Resume Next
    For Each oControl In Me.Controls
        With oControl
            .BackColor = eMACROColour.emcBackground
            .Font = MACRO_DE_FONT
            .FontSize = 8
'   TA  08/08/2005: setting of Font.Charset = 1 to allow eastern european characters. CBD2591.
            .Font.Charset = 1
            .ForeColor = eMACROColour.emdEnabledText
            
            Select Case LCase(oControl.Name)
            Case "lblerrortype"
                .BackColor = eMACROColour.emcNonWhiteBackGround
                .Width = Me.Width
            Case "lcmdmore", "lcmdclose"
                .FontUnderline = True
                .MouseIcon = frmImages.CursorHandPoint.MouseIcon
                .MousePointer = 99 'custom
                .ForeColor = eMACROColour.emcLinkText
            Case "lcmdoverrule"
                .FontUnderline = True
                .MouseIcon = frmImages.CursorHandPoint.MouseIcon
                .MousePointer = 99 'custom
            End Select
        End With
    Next
    Err.Clear
    On Error GoTo 0

End Sub

'---------------------------------------------------------------------
Private Sub lstReason_Click()
'---------------------------------------------------------------------
' They chose a reason from the predefineds
'---------------------------------------------------------------------
    
    txtReason.Text = lstReason.Text

End Sub

'---------------------------------------------------------------------
Private Sub txtReason_Change()
'---------------------------------------------------------------------
' Validate what's been entered
'---------------------------------------------------------------------
Dim sReason As String

    If mbLoading Then Exit Sub
    
    sReason = Trim(txtReason.Text)
    If Not gblnValidString(sReason, valOnlySingleQuotes) Then
        DialogInformation "Overrule reason text" & gsCANNOT_CONTAIN_INVALID_CHARS
    Else
        msOverruleReason = sReason
        mbChanged = True
    End If


    If mnStatus = eStatus.OKWarning Then
        If msOverruleReason = "" Then
            imgIcon.Picture = frmImages.imglistStatus.ListImages(DM30_ICON_WARNING).Picture
        Else
            imgIcon.Picture = frmImages.imglistStatus.ListImages(DM30_ICON_OK_WARNING).Picture
        End If
    End If
    
    
End Sub

'---------------------------------------------------------------------
Private Sub txtReason_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
' NCJ 6 Jan 03 - Copied from frmInputBox.frm
' Validate character entered
'---------------------------------------------------------------------
    
    If Not gblnValidString(Chr(KeyAscii), valOnlySingleQuotes) Then
        KeyAscii = 0
    End If
    'ASH 10/07/2002 CBB 2.2.19 no.20 do not allow enter key in singleline mode
    If KeyAscii = 13 Then
        KeyAscii = 0
    End If

End Sub
'----------------------------------------------------------------------------------------'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------------------'
'unload form if OK or cancel pressed
'when not in warning/ warn overrule
'----------------------------------------------------------------------------------------'

    If lcmdClose = "Close" Then
        Select Case KeyCode
        Case vbKeyEscape, vbKeyReturn
            Unload Me
        End Select
    End If
    
End Sub
