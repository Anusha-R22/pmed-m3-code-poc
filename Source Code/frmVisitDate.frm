VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmVisitDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visit - "
   ClientHeight    =   1770
   ClientLeft      =   12150
   ClientTop       =   2475
   ClientWidth     =   2610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   2610
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDate 
      Height          =   375
      Left            =   660
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1380
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   4095
      Left            =   0
      TabIndex        =   3
      Top             =   1920
      Width           =   5655
      _Version        =   524288
      _ExtentX        =   9975
      _ExtentY        =   7223
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2000
      Month           =   3
      Day             =   20
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   0   'False
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblformat 
      Alignment       =   2  'Center
      Caption         =   "format"
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   960
      Width           =   2475
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Caption         =   "Please enter the visit date"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   120
      Width           =   2475
   End
End
Attribute VB_Name = "frmVisitDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000-2001. All Rights Reserved
'   File:       frmVisitDate.frm
'   Author:     Toby Aldridge March 2000
'   Purpose:    Allows user to enter a visit date
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   NCJ 21 Sep 01 - Updated for MACRO 2.2 (new Arezzo calls etc.)
'                   Also may be used for Form dates too
'
'----------------------------------------------------------------------------------------'

Option Explicit

' The date format we'll use
Private msDateFormat As String

' true if OK clicked
Private mbOK As Boolean
Private mdblNewDate As Double

' NCJ 21 Sep 01 - Remove this dummy function when frmStudyVisits has gone
Public Function Display(s As String, d As Double) As Double

End Function

'---------------------------------------------------------------------
Public Function AskForDate(ByVal sVisitFormName As String, _
                        ByVal dblDate As Double, _
                        ByVal bIsVisit As Boolean) As Double
'---------------------------------------------------------------------
'   Display form and return user's input date
'   Input:
'       sVisitFormName - Visit or eForm name
'       dblVisit - Visit/eForm Date (0 for uninitialised)
'       bIsVisit - TRUE for Visit, FALSE for eForm
'   Output:
'       function - New visit/form date as Double (may be 0)
'---------------------------------------------------------------------

    mbOK = False
    
    ' Use Study's default date format
    msDateFormat = goStudyDef.DateFormat
    
    mdblNewDate = dblDate
    Load Me
    'set caption
    If bIsVisit Then
        Me.Caption = "Visit - " & sVisitFormName
        lblDate.Caption = "Please enter the visit date:"
    Else
        Me.Caption = "eForm - " & sVisitFormName
        lblDate.Caption = "Please enter the eForm date:"
    End If
    lblFormat.Caption = "(" & msDateFormat & ")"
    
    If mdblNewDate = 0 Then
        ' Uninitialised date
        txtDate.Text = ""
    Else
        ' Valid date passed in so put in textbox
        txtDate.Text = Format(CDate(mdblNewDate), msDateFormat)
        txtDate.SelStart = Len(txtDate.Text)
    End If

    'ensure mouse pointer is default
    HourglassSuspend
    FormCentre Me
    'show form
    Me.Show vbModal
    
    HourglassResume
   
    If mbOK Then
        'OK pressed
        AskForDate = mdblNewDate
    Else
        ' Return what we started with
        AskForDate = dblDate
    End If
    
    Exit Function
    
ErrHandler:

    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "AskForDate")
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
Private Sub cmdCancel_Click()
'---------------------------------------------------------------------
' User presses Cancel button
'---------------------------------------------------------------------
    
    Unload Me

End Sub

'---------------------------------------------------------------------
Private Sub cmdOK_Click()
'---------------------------------------------------------------------
' Validate user's input and display error message if appropriate
' Store "double" date value if OK
' Assume we can't get here if txtDate is empty
'---------------------------------------------------------------------
Dim sDate As String
Dim sMsg As String
Dim dblDate As Double

    On Error GoTo ErrHandler

    sDate = Trim(txtDate.Text)  ' Should not be ""
    sMsg = ValidateDate(sDate, dblDate)
    If sMsg > "" Then
        ' It was not accepted
        Call DialogInformation(sMsg)
        txtDate.SetFocus
    Else
        ' It was OK
        ' store Double value
        mdblNewDate = dblDate
        ' OK pressed
        mbOK = True
        ' Unload form
        Unload Me
    End If
    
Exit Sub
    
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdOK_Click")
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
Private Function ValidateDate(ByVal sInDate As String, ByRef dblDate As Double) As String
'---------------------------------------------------------------------
' Validate text date value
' Returns empty string if all OK, otherwise returns error message
' dblDate is set to the date as a double (if valid)
' Assume sDate > ""
'---------------------------------------------------------------------
Dim sMsg As String
Dim sArezzoDate As String
Dim sDate As String

    sMsg = ""
    dblDate = 0
    ' Read date using current default date format
    sDate = goArezzo.ReadValidDate(sInDate, msDateFormat, sArezzoDate)
    If sDate = "" Then
        ' empty string, therefore invalid
        sMsg = sInDate & " is not recognised as a valid date." & vbCrLf
        sMsg = sMsg & "Please enter the date in the format " & msDateFormat
    Else
        ' It was a valid date - check it's reasonable
        dblDate = goArezzo.ArezzoDateToDouble(sArezzoDate)
        If dblDate <= 0 Then
            ' It's too far in the past
            sMsg = sDate & " is not accepted as a valid date." & vbCrLf
            sMsg = sMsg & "The date must not be before 1900"
            dblDate = 0
        ElseIf dblDate > CDbl(Now) Then
            ' It's in the future
            sMsg = sDate & " is not accepted as a valid date." & vbCrLf
            sMsg = sMsg & "The date must not be in the future."
            dblDate = 0
        Else
            ' It was OK
        End If
    End If
        
    ValidateDate = sMsg

End Function

Private Sub Form_Load()
    Me.Icon = frmMenu.Icon
End Sub

'---------------------------------------------------------------------
Private Sub txtDate_Change()
'---------------------------------------------------------------------
' Disable OK if no date entered
'---------------------------------------------------------------------

    cmdOK.Enabled = (Trim(txtDate.Text) > "")

End Sub

'---------------------------------------------------------------------
Private Sub txtDate_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
' Interpret RETURN as OK button
' Show today's date if "t" has been entered
'---------------------------------------------------------------------

    If KeyAscii = Asc(vbCr) Then
        ' Convert "t" to today's date
        If LCase(Trim(txtDate.Text)) = "t" Then
            txtDate.Text = Format(Now, msDateFormat)
        End If
        Call cmdOK_Click
    End If

End Sub
