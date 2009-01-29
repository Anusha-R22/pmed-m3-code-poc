VERSION 5.00
Begin VB.Form frmOptionMsgBox 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1635
   ClientLeft      =   6075
   ClientTop       =   3000
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3420
      TabIndex        =   3
      Top             =   540
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3420
      TabIndex        =   2
      Top             =   120
      Width           =   1155
   End
   Begin VB.Frame fraOption 
      Height          =   1575
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   3255
      Begin VB.OptionButton optOption 
         Caption         =   "Option1"
         Height          =   315
         Index           =   0
         Left            =   540
         TabIndex        =   0
         Top             =   780
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblOption 
         AutoSize        =   -1  'True
         Caption         =   "lblOption"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   360
         Width           =   2895
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmOptionMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   File:       frmOptionMsgBox.frm
'   Author:     Toby Aldridge March 2000
'   Purpose:    Allows user to choose from an options list
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'


Option Explicit

Private Const msPrefix = "opt"

Private mlOption As Long
Private mbStatus As Boolean

'---------------------------------------------------------------------
Public Function Display(sCaption As String, sFrame As String, sLabel As String, _
                        sOption As String, Optional sOK As String, Optional sCancel As String, _
                        Optional bOK As Boolean = True, Optional bCancel As Boolean = True) As Long
'---------------------------------------------------------------------
'TA 29/03/2000: Display Form
'   Input:
'       sCaption - form title
'       sFrame - frame caption
'       sOption - string containg list of options separated by "|"
'   Output:
'       function - number in list of option selected / gsMINUS_ONE if cancelled
'---------------------------------------------------------------------
       
    On Error GoTo ErrHandler
           
    mbStatus = False
    
    Me.Caption = sCaption
    If sOK <> "" Then
        cmdOK.Caption = sOK
    End If
    
    If sCancel <> "" Then
        cmdCancel.Caption = sCancel
    End If
    
    cmdOK.Visible = bOK
    cmdCancel.Visible = bCancel
    
    fraOption.Caption = sFrame
    lblOption.Width = Me.TextWidth(sLabel & "00000")
    lblOption.Caption = sLabel
    
    HourglassSuspend
    
    OptionShow sOption

    FormCentre Me

    Me.Show vbModal
    
    If mbStatus Then
        Display = mlOption
    Else
        Display = glMINUS_ONE
    End If
    
    HourglassResume
    
    Exit Function
    
ErrHandler:

    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "Display")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Function

'---------------------------------------------------------------------
Private Sub OptionShow(sOption As String)
'---------------------------------------------------------------------
'TA 29/03/2000: Display Form
'add controls to form at run time
'---------------------------------------------------------------------

Dim i As Long
Dim vOption As Variant
Dim lTop As Long
Dim lWidth As Long
Dim lHeight As Long
Dim lLeft As Long
Dim lTitleHeight As Long

    On Error GoTo ErrHandler

    vOption = Split(sOption, "|")
    lTitleHeight = Me.Height - Me.ScaleHeight
    lHeight = 255
    lLeft = 480
    lTop = lblOption.Top + lblOption.Height + 240
    lWidth = lblOption.Width + 480
    
    For i = 0 To UBound(vOption)
        Load Me.optOption(i + 1)
        Set optOption(i + 1).Container = fraOption
        With optOption(i + 1)
            .Visible = True
            .Top = lTop
            If Me.TextWidth(vOption(i) & "0000") > lWidth Then
                lWidth = Me.TextWidth(vOption(i) & "0000")
            End If
            .Left = lLeft
            .Width = lWidth
            .Height = lHeight
            .Caption = vOption(i)
            .TabIndex = i
            .TabStop = True
        End With
        lTop = lTop + lHeight + 60
    Next
    
    cmdOK.TabIndex = UBound(vOption) + 1
    cmdCancel.TabIndex = UBound(vOption) + 2
    
    fraOption.Width = lLeft + lWidth + 120
    cmdOK.Left = fraOption.Left + fraOption.Width + 120
    cmdCancel.Left = cmdOK.Left
    
    fraOption.Height = lTop + 120
    Me.Width = cmdOK.Left + cmdOK.Width + 180
    Me.Height = fraOption.Height + 60 + lTitleHeight
    
    
    Exit Sub
    
ErrHandler:

    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "OptionShow")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdCancel_Click()
'---------------------------------------------------------------------
    Unload Me
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdOK_Click()
'---------------------------------------------------------------------
'TA 29/03/2000: Check which option is selected and exit
'---------------------------------------------------------------------
    
    mbStatus = True
    Unload Me

End Sub

Private Sub Form_Load()
    Me.Icon = frmMenu.Icon
End Sub

'---------------------------------------------------------------------
Private Sub optOption_Click(Index As Integer)
'---------------------------------------------------------------------

    mlOption = Index

End Sub

'---------------------------------------------------------------------
Private Sub optOption_LostFocus(Index As Integer)
'---------------------------------------------------------------------
    optOption(Index).TabStop = True

End Sub
