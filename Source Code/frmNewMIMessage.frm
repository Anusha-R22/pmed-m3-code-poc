VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmNewMIMessage 
   Caption         =   "NB status frame is hidden off to the right"
   ClientHeight    =   3480
   ClientLeft      =   4050
   ClientTop       =   5745
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   6750
   Begin VB.CommandButton cmdExternal 
      Caption         =   "Copy from OC form"
      Height          =   315
      Left            =   4140
      TabIndex        =   22
      Top             =   2580
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame fraDiscrepancy 
      Height          =   795
      Left            =   60
      TabIndex        =   15
      Top             =   2580
      Width           =   3975
      Begin MSComCtl2.UpDown spnPriority 
         Height          =   345
         Left            =   3600
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtOCId 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1560
         TabIndex        =   1
         Top             =   240
         Width           =   945
      End
      Begin VB.TextBox txtPriority 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   3180
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "5"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Priority"
         Height          =   270
         Index           =   3
         Left            =   2640
         TabIndex        =   17
         Top             =   300
         Width           =   480
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "OC Discrepancy Id"
         Height          =   270
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.Frame fraNote 
      Caption         =   "Status"
      Height          =   735
      Left            =   10560
      TabIndex        =   6
      Top             =   2580
      Width           =   3255
      Begin VB.OptionButton optPublic 
         Caption         =   "Public"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   300
         Width           =   1095
      End
      Begin VB.OptionButton optPrivate 
         Caption         =   "Private"
         Height          =   195
         Left            =   1440
         TabIndex        =   19
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.Frame fraSDV 
      Caption         =   "Status"
      Height          =   735
      Left            =   6960
      TabIndex        =   9
      Top             =   2040
      Width           =   3255
      Begin VB.OptionButton optQueried 
         Caption         =   "Queried"
         Height          =   195
         Left            =   1200
         TabIndex        =   20
         Top             =   300
         Width           =   915
      End
      Begin VB.OptionButton optDone 
         Caption         =   "Done"
         Height          =   195
         Left            =   2220
         TabIndex        =   4
         Top             =   300
         Width           =   915
      End
      Begin VB.OptionButton optPlanned 
         Caption         =   "Planned"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.Frame fraMessage 
      Caption         =   "Please enter message text"
      Height          =   1695
      Left            =   60
      TabIndex        =   14
      Top             =   840
      Width           =   6615
      Begin VB.TextBox txtMessage 
         Height          =   1335
         Left            =   120
         MaxLength       =   2000
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   6375
      End
   End
   Begin VB.Frame fraQuestion 
      Caption         =   "Question details"
      Height          =   735
      Left            =   60
      TabIndex        =   10
      Top             =   60
      Width           =   6615
      Begin VB.TextBox txtQuestion 
         BackColor       =   &H8000000F&
         Height          =   345
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   2025
      End
      Begin VB.TextBox txtValue 
         BackColor       =   &H8000000F&
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   2025
      End
      Begin VB.Label lblValue 
         Caption         =   "Value"
         Height          =   195
         Left            =   3840
         TabIndex        =   21
         Top             =   300
         Width           =   495
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         Caption         =   "Question name"
         Height          =   270
         Left            =   240
         TabIndex        =   13
         Top             =   300
         Width           =   1140
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4140
      TabIndex        =   7
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5460
      TabIndex        =   8
      Top             =   3000
      Width           =   1215
   End
End
Attribute VB_Name = "frmNewMIMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   File:       frmNewMIMessage.frm
'   Author:     Toby Aldridge May 2000
'   Purpose:    Let User enter discrepancy
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   TA 31/05/2000: Amendments as in document 'Biderectional Communication Code review' - Will Casey, 30/05/2000
'   TA 12/06/2000 SR3586: Locked property of the value text box is now set to true to prevent editing
' TA 03/04/02 - Added OC discrepancy code done in Palo Alto
' ASH/ZA 10/07/2002 - CBB 2.2.15 no.R24 Added Function ValidateTextField.
' NCJ 14 Oct 02 - Added new stuff for new SDV statuses and Scope
' TA 23/01/2003: allow up to 2000 characters for text- bug 673 in 3.0 issues and bugs db
'----------------------------------------------------------------------------------------'

Option Explicit

Private mbOK As Boolean

Private msMsgText As String
Private mlOCId As Long
Private mnPriority As Integer
Private mnStatus As Integer

'za 08/07/2002 - added this variable to hold message type
Private msMessageType As String
Private msStudyName As String


'----------------------------------------------------------------------------------------'
Public Function Display(nMIMsgType As MIMsgType, _
                        enScope As MIMsgScope, _
                        sStudy As String, sObjName As String, sValue As String, _
                        sMsgText As String, lOCId As Long, _
                        nPriority As Integer, nStatus As Integer) As Boolean
'----------------------------------------------------------------------------------------'
'   Display form
'   Input:
'       nMIMsgType - MIMessage type
'       enScope - MIMessage scope
'       sStudy - study name
'       sObjName - visit/eform/question name or Subject Id/Label
'       sValue - value of question (if scope = question, ignored otherwise)
'   Output:
'       sMsgText - message text entered
'       lOCId - Oracle Clincical Id entered (discrepancy only)
'       nPriority - priorty entered (discrepancy only)
'       nStatus - status entered (SDV only)
'       function - OK Clicked?
'----------------------------------------------------------------------------------------'
' NCJ 14 Oct 02 - Added enScope argument
'----------------------------------------------------------------------------------------'

    mbOK = False
    
    Load Me
    Me.Icon = frmMenu.Icon
        
    fraMessage.Caption = "Enter " & GetMIMTypeText(nMIMsgType) & " text"
    
    'za 08/07/2002 - added this variable to hold message type
    msMessageType = GetMIMTypeText(nMIMsgType)
    
    msStudyName = sStudy
    
    Select Case nMIMsgType
    Case MIMsgType.mimtDiscrepancy
        fraSDV.Visible = False
        fraNote.Visible = False
        fraDiscrepancy.Visible = True
        If Not frmMenu.gOC Is Nothing Then
            cmdExternal.Visible = frmMenu.gOC.HaveInfo
        End If
    Case MIMsgType.mimtSDVMark
        fraDiscrepancy.Visible = False
        fraNote.Visible = False
        fraSDV.Visible = True
        fraSDV.Left = fraDiscrepancy.Left
        fraSDV.Top = fraDiscrepancy.Top
        optPlanned.Value = vbChecked
    Case MIMsgType.mimtNote
        fraDiscrepancy.Visible = False
        fraSDV.Visible = False
        fraNote.Left = fraDiscrepancy.Left
        optPublic.Value = vbChecked
    End Select
    
    txtValue.Text = sValue
    ' Ony show Value fields if scope is Question
    lblValue.Visible = (enScope = MIMsgScope.mimscQuestion)
    txtValue.Visible = (enScope = MIMsgScope.mimscQuestion)

    ' Set labels to show correct scope
    lblName.Caption = MACROMIMsgBS30.GetScopeText(enScope)
    If enScope = MIMsgScope.mimscSubject Then
        lblName.Caption = lblName.Caption & " label"
    Else
        lblName.Caption = lblName.Caption & " name"
    End If
    fraQuestion.Caption = MACROMIMsgBS30.GetScopeText(enScope) & " details"
    
    Me.Caption = "New " & GetMIMTypeText(nMIMsgType) & " - " & sStudy
    txtQuestion.Text = sObjName
    
    FormCentre Me
    Me.Show vbModal
    
    If mbOK Then
        sMsgText = msMsgText
        lOCId = mlOCId
        nPriority = mnPriority
        nStatus = mnStatus
        If Not frmMenu.gOC Is Nothing Then
            frmMenu.gOC.CheckItem lOCId
        End If
    End If
    
    Display = mbOK
    
Exit Function
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "Display")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Function

'----------------------------------------------------------------------------------------'
Private Sub cmdCancel_Click()
'----------------------------------------------------------------------------------------'
    Unload Me
End Sub

'----------------------------------------------------------------------------------------'
Private Sub cmdExternal_Click()
'----------------------------------------------------------------------------------------'
'paste in info from OC
'----------------------------------------------------------------------------------------'

    txtMessage.Text = frmMenu.gOC.DiscrepancyText
    txtOCId.Text = frmMenu.gOC.OCId

End Sub

'----------------------------------------------------------------------------------------'
Private Sub cmdOK_Click()
'----------------------------------------------------------------------------------------'
    
    If CheckFields Then
    
        msMsgText = Trim(txtMessage.Text)
        If fraDiscrepancy.Visible Then
            mlOCId = CLng(Val(txtOCId.Text))
            mnPriority = Val(txtPriority.Text)
        ElseIf fraSDV.Visible Then
            If optPlanned.Value Then
                mnStatus = eSDVMIMStatus.ssPlanned
            ElseIf optDone.Value Then
                mnStatus = eSDVMIMStatus.ssDone
            ElseIf optQueried.Value Then        ' NCJ 14 Oct 02
                mnStatus = eSDVMIMStatus.ssQueried
            End If
        Else
            If optPublic.Value Then
                mnStatus = eNoteMIMStatus.nsPublic
            ElseIf optPrivate.Value Then
                mnStatus = eNoteMIMStatus.nsPrivate
            End If
   
        End If

        mbOK = True
        Unload Me
    End If

End Sub

'----------------------------------------------------------------------------------------'
Private Function CheckFields() As Boolean
'----------------------------------------------------------------------------------------'
' validate entered values if this is a discrepancy
'----------------------------------------------------------------------------------------'
Dim bValidated As Boolean
Dim lOCId As Long

    bValidated = True
    
    If fraDiscrepancy.Visible Then
        txtOCId.Text = Trim(txtOCId.Text)
        txtPriority.Text = Trim(txtPriority.Text)
        
        If Not (txtOCId.Text = "" Or IsNumeric(txtOCId.Text)) Then
            MsgBox "The O.C. Discrepancy must be a positive number", vbOKOnly + vbCritical, "Discrepancy Error"
            bValidated = False
        Else
            'if too big for a long
            On Err GoTo ErrNotLong
            lOCId = CLng(Val(txtOCId.Text))
            On Error GoTo ErrHandler
            If lOCId <> 0 Then
                'check oc id does not already exist
                If MACROMIMsgBS30.IsDuplicateOCId(gsADOConnectString, lOCId, msStudyName) Then
                    'already exists in this study - exclude
                    MsgBox "The O.C. Discrepancy id alredy exists", vbOKOnly + vbCritical, "Discrepancy Error"
                    bValidated = False
                End If
            End If
        End If
    
        If Not ((Val(txtPriority.Text) > 0) And (Val(txtPriority.Text) < 11)) Then
            MsgBox "The priority must be a number from 1 to 10", vbOKOnly + vbCritical, "Discrepancy Error"
            bValidated = False
        End If
    End If
    
    'ASH 10/07/2002 CBB 2.2.15 no.R24
    'Checks for invalid characters either typed or pasted
     If Not ValidateTextField(txtMessage.Text) Then
        bValidated = False
    End If
    
    CheckFields = bValidated

Exit Function
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "CheckFields")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
Exit Function

ErrNotLong:

    MsgBox "The O.C. Discrepancy id is not a valid value", vbOKOnly + vbCritical, "Discrepancy Error"
    CheckFields = False
    

End Function


'----------------------------------------------------------------------------------------'
Private Sub spnPriority_DownClick()
'----------------------------------------------------------------------------------------'
Dim nPriority As Integer
    
    nPriority = Val(txtPriority.Text)
    If nPriority > 1 Then
        txtPriority.Text = Format(nPriority - 1)
    End If

End Sub

'----------------------------------------------------------------------------------------'
Private Sub spnPriority_UpClick()
'----------------------------------------------------------------------------------------'
Dim nPriority As Integer
    
    nPriority = Val(txtPriority.Text)
    If nPriority < 10 Then
        txtPriority.Text = Format(nPriority + 1)
    End If

End Sub

'----------------------------------------------------------------------------------------'
Private Sub txtMessage_KeyPress(KeyAscii As Integer)
'----------------------------------------------------------------------------------------'

    If KeyAscii <> Asc(vbCr) And KeyAscii <> Asc(vbLf) Then
        If Not gblnValidString(Chr(KeyAscii), valOnlySingleQuotes) Then
            KeyAscii = 0
        End If
    End If
    
End Sub

'------------------------------------------------------------------------------------------
Private Function ValidateTextField(ByVal sString As String) As Boolean
'------------------------------------------------------------------------------------------
'ASH 10/07/2002 CBB 2.2.15 no.R24
'Checks if text string does not contain invalid characters such as |,",` etc.
'------------------------------------------------------------------------------------------
On Error GoTo ErrLabel

    ValidateTextField = True

    If Not gblnValidString(sString, valOnlySingleQuotes) Then
        DialogInformation msMessageType & " text" & gsCANNOT_CONTAIN_INVALID_CHARS
        ValidateTextField = False
        'TA 23/01/2003: allow up to 2000 characters - bug 673 in 3.0 issues and bugs db
    ElseIf Len(sString) > 2000 Then
        DialogInformation msMessageType & " text may not be more than 2000 characters"
        ValidateTextField = False
    End If

Exit Function
ErrLabel:
  Err.Raise Err.Number, , Err.Description & "|" & "frmNewMIMessage.ValidateTextField"
End Function
