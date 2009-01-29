VERSION 5.00
Begin VB.Form frmQuestionDetails 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   2400
   ClientLeft      =   1740
   ClientTop       =   1755
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtComment 
      Height          =   1395
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   840
      Width           =   6315
   End
   Begin VB.Label lblQuestion 
      Caption         =   "lblQuestion"
      Height          =   240
      Left            =   0
      TabIndex        =   2
      Top             =   15
      Width           =   4755
   End
   Begin VB.Label lcmdClose 
      Caption         =   "Close"
      Height          =   240
      Left            =   5760
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   0
      Width           =   675
   End
   Begin VB.Image img 
      Height          =   315
      Index           =   0
      Left            =   2220
      Top             =   360
      Width           =   315
   End
   Begin VB.Label label 
      Caption         =   "dummy"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frmQuestionDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000-2002. All Rights Reserved
'   File:       frmQuestionDetails.frm
'   Author:     Toby Aldridge December  2002
'   Purpose:    Display question details in a modal form
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
' Revisions:
'   NCJ 17 Feb 03 - Added RFC; removed "dummy text; made sure long values show
'   NCJ 28 Mar 03 - Make sure ampersands handled correctly (by doubling them up in label captions)
'   TA 05/11/2004: BD2365, CRM 710 autosize the response value field height to display whole response
'   TA  1/7/2005    set font.charset to allow non western european characters
'----------------------------------------------------------------------------------------'


Option Explicit


Private moResponse As Response


Private Enum eLabel
    
    lC_Value = 1
    lValue = 2
    lC_status = 3
    lStatus = 4
    lStatusMsg = 5
    lC_Lock = 6
    lLock = 7
    lC_Disc = 8
    lDisc = 9
    lC_SDV = 10
    lSDV = 11
    lc_NR = 12
    lNR = 13
    lC_Notes = 14
    lNotes = 15
    lAuth = 16
    lC_RFC = 17     ' NCJ 17 Feb 03
    lRFC = 18
    lC_Comment = 19
    lComment = 20
'    lC_Comment = 17
'    lComment = 18
    
End Enum

Private Const mCAPTION_WIDTH = 1300
Private Const mCAPTION_HEIGHT = 315
Private Const mSTATUS_MSG_HEIGHT = 1300
Private Const mCOMMENT_HEIGHT = 1300
Private Const mGAP_WIDTH = 60
Private Const mGAP_HEIGHT = 20

Private moUser As MACROUser

'----------------------------------------------------------------------------------------'
Public Sub Display(oResponse As Response, oCoord As clsCoords, oUser As MACROUser)
'----------------------------------------------------------------------------------------'

    
    Me.Top = oCoord.Row
    Me.Left = oCoord.Col
    Set moResponse = oResponse
    Set moUser = oUser
    Call UnloadControls
    Call SetUpLabels
    
    'ensure we appear completely on the screen
    If Me.Top < 100 Then
        Me.Top = 100
    End If
    If Me.Top + Me.Height > Screen.Height - 500 Then
        Me.Top = Screen.Height - Me.Height - 500
    End If
    If Me.Left < 100 Then
        Me.Left = 100
    End If
    If Me.Left + Me.Width > Screen.Width Then
        Me.Left = Screen.Width - Me.Width
    End If
        
    Me.Show vbModal

End Sub

'----------------------------------------------------------------------------------------'
Private Sub UnloadControls()
'----------------------------------------------------------------------------------------'
'unload any previously load controls
'----------------------------------------------------------------------------------------'
Dim oCon As Control

    For Each oCon In Me.Controls
        If oCon.Name = "label" Or oCon.Name = "img" Then
            If oCon.Index <> 0 Then
                Unload oCon
            End If
        End If
    Next
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub SetUpLabels()
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'
Dim lTop As Long
Dim enLabel As eLabel
Dim oCon As Control

    Me.BackColor = emcBackground
    
    On Error Resume Next
    For Each oCon In Me.Controls
        With oCon
            .BackColor = eMACROColour.emcBackground
            .Font = MACRO_DE_FONT
            .Font.Charset = 1
            .FontSize = 8
            .ForeColor = eMACROColour.emdEnabledText
            
            Select Case LCase(oCon.Name)
            Case "lcmdclose"
                lcmdClose.ForeColor = eMACROColour.emcLinkText
                .FontUnderline = True
                .MouseIcon = frmImages.CursorHandPoint.MouseIcon
                .MousePointer = 99 'custom
            End Select
        End With
    Next
    Err.Clear
    On Error GoTo 0
    
    lblQuestion.Caption = "Question: " & QuestionNameText(moResponse)
    lblQuestion.FontBold = True
    
    With label(0)
        .Visible = False
        .Caption = ""   ' Avoid it saying "dummy"!
        
    End With

    lTop = lblQuestion.Top + lblQuestion.Height + mGAP_HEIGHT + 240
    
    For enLabel = lC_Value To lComment
        Load Me.label(enLabel)
        With Me.label(enLabel)
            .Visible = True
            .Font = MACRO_DE_FONT
            .FontSize = MACRO_DE_FONT_SIZE
            .BackColor = Me.BackColor
            .Top = lTop
            .Left = mGAP_WIDTH
            .Width = mCAPTION_WIDTH
            .Height = mCAPTION_HEIGHT
        
            Select Case enLabel
            'label specific stuff
            Case lC_Value
                .Caption = "Value"
            Case lValue
                'TA 05/11/2004: BD2365, CRM 710 autosize the response value field height to display whole response
                .WordWrap = True
                .Caption = moResponse.Value
                .AutoSize = True
                .Left = mCAPTION_WIDTH + (2 * mGAP_WIDTH)
                .Width = Me.Width - .Left - 120
                lTop = lTop + .Height + mGAP_HEIGHT
            Case lC_status
                .Caption = "Status"
            Case lStatus
                .Caption = moResponse.StatusString
                .Left = mCAPTION_WIDTH + (2 * mGAP_WIDTH)
                .Height = mCAPTION_HEIGHT * 2
                Load Me.img(enLabel)
                Me.img(enLabel).Left = .Left
                Me.img(enLabel).Top = .Top
                'show icon
                GetEFormBuilder.SetBaseStatusImage moResponse.Status, Me.img(enLabel)
                Me.img(enLabel).Visible = True
                'move label over
                .Left = Me.img(enLabel).Left + Me.img(enLabel).Width + mGAP_WIDTH
            Case lStatusMsg
                .WordWrap = True
                .Height = mSTATUS_MSG_HEIGHT
                .Caption = Replace(moResponse.ValidationMessage, "&", "&&")
                .Width = 4 * mCAPTION_WIDTH
                If moResponse.Status = eStatus.OKWarning Then
                    .Caption = .Caption & vbCrLf & vbCrLf & "Reason for overrule:" & vbCrLf & Replace(moResponse.OverruleReason, "&", "&&")
                End If
                lTop = lTop + .Height + mGAP_HEIGHT
                .Left = label(lStatus).Left + label(lStatus).Width + mGAP_WIDTH
                .Width = Me.ScaleWidth - .Left - 60
                
            Case lC_Lock
                .Caption = "Lock status"
            Case lLock
                .Caption = moResponse.LockStatusString
                lTop = lTop + .Height + mGAP_HEIGHT
                .Left = mCAPTION_WIDTH + (2 * mGAP_WIDTH)
                Load Me.img(enLabel)

                Me.img(enLabel).Left = .Left
                Me.img(enLabel).Top = .Top
                'show icon
                GetEFormBuilder.SetLockStatusImage Me.img(enLabel), moResponse.LockStatus
                Me.img(enLabel).Visible = True
                'move label over
                .Left = Me.img(enLabel).Left + Me.img(enLabel).Width + mGAP_WIDTH
            Case lC_Disc
                .Caption = "Discrepancies"
            Case lDisc
                .Caption = GetHeirachicalMIMsgText(MIMsgType.mimtDiscrepancy, moResponse.DiscrepancyStatus)
                lTop = lTop + .Height + mGAP_HEIGHT
                .Left = mCAPTION_WIDTH + (2 * mGAP_WIDTH)
                Load Me.img(enLabel)

                Me.img(enLabel).Left = .Left
                Me.img(enLabel).Top = .Top

                'show icon
                GetEFormBuilder.SetDiscStatusImage Me.img(enLabel), moResponse.DiscrepancyStatus
                Me.img(enLabel).Visible = True
                'move label over
                .Left = Me.img(enLabel).Left + Me.img(enLabel).Width + mGAP_WIDTH
            Case lC_SDV
                .Caption = "SDV"
            Case lSDV
                .Caption = GetHeirachicalMIMsgText(MIMsgType.mimtSDVMark, moResponse.SDVStatus)
                lTop = lTop + .Height + mGAP_HEIGHT
                .Left = mCAPTION_WIDTH + (2 * mGAP_WIDTH)
                Load Me.img(enLabel)

                Me.img(enLabel).Left = .Left
                Me.img(enLabel).Top = .Top + (Me.img(enLabel).Height / 2 - 20)
                Me.img(enLabel).Height = 40
                
                'show icon
                GetEFormBuilder.SetSDVStatusImage moResponse.SDVStatus, Me.img(enLabel)
                Me.img(enLabel).Visible = True
                'move label over
                .Left = Me.img(enLabel).Left + Me.img(enLabel).Width + mGAP_WIDTH
            Case lc_NR
                .Caption = "NR Status"
                 .Visible = (moResponse.Element.DataType = eDataType.LabTest)
                 
             Case lNR
                .Visible = label(lc_NR).Visible
                If .Visible Then
                    .Caption = GetNRCTCText(moResponse.NRStatus, moResponse.CTCGrade)
                    lTop = lTop + .Height + mGAP_HEIGHT
                    .Left = mCAPTION_WIDTH + (2 * mGAP_WIDTH)
                End If
            Case lC_Notes
                .Caption = "Notes"
            Case lNotes
                .Caption = IIf(moResponse.NoteStatus = nsNoNote, "None", "Yes")
                .Left = mCAPTION_WIDTH + (2 * mGAP_WIDTH)
                lTop = lTop + .Height + mGAP_HEIGHT
            Case lAuth
                ' Authorisation role
                If moResponse.Element.Authorisation > "" Then
                    .Width = Me.Width
                    .Caption = "This question needs to be authorised by a user with role: " & moResponse.Element.Authorisation
                    lTop = lTop + .Height + mGAP_HEIGHT
                Else
                    .Visible = False
                End If
            ' NCJ 17 Feb 03 - Added RFC
            Case lC_RFC
                .Caption = "RFC"
                .Visible = (moResponse.ReasonForChange > "")
            Case lRFC
                ' Reason for Change
                If moResponse.ReasonForChange > "" Then
                    .WordWrap = True
                    .Height = mSTATUS_MSG_HEIGHT
                    .Caption = Replace(moResponse.ReasonForChange, "&", "&&")   ' Deal with &
                    .Left = mCAPTION_WIDTH + (2 * mGAP_WIDTH)
                    lTop = lTop + .Height + mGAP_HEIGHT
                    .Width = Me.ScaleWidth - .Left - 60
                Else
                    .Visible = False
                End If
                
            Case lC_Comment
                .Caption = "Comments"
                .Left = mGAP_WIDTH
                 lTop = lTop + .Height + mGAP_HEIGHT
                .Visible = (moResponse.Comments > "") And (goUser.CheckPermission(gsFnViewIComments))
            End Select
            
        
        End With
    Next
    
    
    If label(lC_Comment).Visible Then
        With txtComment
            .Visible = True
            .Font = MACRO_DE_FONT
            .FontSize = MACRO_DE_FONT_SIZE
            .BackColor = Me.BackColor
            .Top = lTop
            .Left = mGAP_WIDTH
            .Width = Me.Width - (2 * mGAP_WIDTH) - 120
            .Locked = True
            txtComment.Top = lTop
            txtComment.Text = moResponse.Comments
        End With
        Me.Height = (Me.Height - Me.ScaleHeight) + txtComment.Top + txtComment.Height + 120
    Else
        txtComment.Visible = False
        Me.Height = (Me.Height - Me.ScaleHeight) + label(lC_Comment).Top
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Private Function GetEFormBuilder() As EFormBuilder
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'
    
    Set GetEFormBuilder = New EFormBuilder
    Set GetEFormBuilder.User = moUser
    
End Function

'----------------------------------------------------------------------------------------'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------------------'
'unload form if OK or cancel pressed
'----------------------------------------------------------------------------------------'

    Select Case KeyCode
    Case vbKeyEscape, vbKeyReturn
        Unload Me
    End Select
    
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Form_Unload(Cancel As Integer)
'----------------------------------------------------------------------------------------'
    
    Set moResponse = Nothing

End Sub

'----------------------------------------------------------------------------------------'
Private Sub lcmdclose_Click()
'----------------------------------------------------------------------------------------'
        
        Unload Me

End Sub
