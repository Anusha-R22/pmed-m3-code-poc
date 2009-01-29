VERSION 5.00
Begin VB.Form frmLocalFormats 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Local Date Format"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   3630
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDateTime 
      Height          =   375
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   21
      TabIndex        =   4
      Text            =   "dd/mm/yyyy hh:mm:ss"
      Top             =   660
      Width           =   1755
   End
   Begin VB.TextBox txtDateFormat 
      Height          =   375
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   180
      Width           =   1755
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   2400
      TabIndex        =   1
      Top             =   1140
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   1140
      TabIndex        =   0
      Top             =   1140
      Width           =   1140
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date/Time format"
      Height          =   315
      Left            =   300
      TabIndex        =   5
      Top             =   705
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Preferred date format"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   240
      Width           =   1635
   End
End
Attribute VB_Name = "frmLocalFormats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       frmLocalFormats.frm
'   Copyright:  InferMed Ltd. 2003-2005. All Rights Reserved
'   Author:     Nicky Johns, January 2003
'   Purpose:    Allow user to edit their preferred local formats
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'Revisions:
'   NCJ 21 Jan 03 - Initial development
'   NCJ 22 Jan 03 - Make non-sizeable; deal with Close box
'   NCJ 8 Dec 05 - Handle new set of date formats (for Partial Dates)
'----------------------------------------------------------------------------------------'

Private mbOK As Boolean

Private moArezzo As Arezzo_DM

Private msOrigDateFormat As String

Private msDateFormat As String

Option Explicit

'----------------------------------------------------------------------------------------'
Public Function Display(ByRef sLocalDate As String, _
                            oArezzo As Arezzo_DM) As Boolean
'----------------------------------------------------------------------------------------'
' Returns TRUE if they did anything, with sLocalDate updated
' Otherwise returns False
'----------------------------------------------------------------------------------------'

    mbOK = False
    
    Set Me.Icon = frmMenu.Icon
    
    Set moArezzo = oArezzo
    txtDateFormat.Text = sLocalDate
    msOrigDateFormat = sLocalDate
    
    FormCentre Me
    
    Me.Show vbModal
    
    ' Return values
    If mbOK Then
        sLocalDate = msDateFormat
    End If
    Display = mbOK
    
    Set moArezzo = Nothing
    
End Function

'----------------------------------------------------------------------------------------'
Private Sub cmdCancel_Click()
'----------------------------------------------------------------------------------------'

    Unload Me
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub cmdOK_Click()
'----------------------------------------------------------------------------------------'

    mbOK = True
    Unload Me
    
End Sub

'----------------------------------------------------------------------------------------'
Private Function ValidateDate() As Boolean
'----------------------------------------------------------------------------------------'
' Returns TRUE if txtDateFormat contains valid date format (allow full date-only)
' and stores it in msDateFormat
' NCJ 8 Dec 05 - Handle new set of date formats
'----------------------------------------------------------------------------------------'
Dim sDate As String
Dim bOK As Boolean

    bOK = False
    
    sDate = Trim(txtDateFormat.Text)
    If gblnValidString(sDate, valOnlySingleQuotes) Then
        ' Ask AREZZO to validate it for us
'        bOK = (moArezzo.ValidateDateFormat(sDate) = eDateTimeType.dttDateOnly)
        Select Case moArezzo.ValidateDateFormat(sDate)
        Case eDateTimeType.dttDMY, eDateTimeType.dttMDY, eDateTimeType.dttYMD
            ' Full dates OK
            bOK = True
        Case Else
            ' Anything else not OK
            bOK = False
        End Select
    End If
    If bOK Then msDateFormat = sDate
    
    ValidateDate = bOK
    
End Function

'----------------------------------------------------------------------------------------'
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'----------------------------------------------------------------------------------------'
' Intercept clicking of Close box
'----------------------------------------------------------------------------------------'

    If UnloadMode = vbFormControlMenu Then
        If msOrigDateFormat <> msDateFormat Then
            If DialogQuestion("Are you sure you want to cancel the dialog and lose your changes?") = vbNo Then
                Cancel = 1
            End If
        End If
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub txtDateFormat_Change()
'----------------------------------------------------------------------------------------'

    EnableOK
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub EnableOK()
'----------------------------------------------------------------------------------------'
' Enable the OK button if fields are valid
'----------------------------------------------------------------------------------------'

    cmdOK.Enabled = ValidateDate
    If cmdOK.Enabled Then
        'Populate combined date/time
       txtDateTime.Text = msDateFormat & " hh:mm:ss"
    End If
    
End Sub

