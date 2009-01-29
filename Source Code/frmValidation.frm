VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmValidation 
   Caption         =   "Validation"
   ClientHeight    =   2430
   ClientLeft      =   8475
   ClientTop       =   7545
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   5145
   Begin VB.Frame fraRules 
      Caption         =   " Validation rules "
      Height          =   1815
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   4995
      Begin MSFlexGridLib.MSFlexGrid grdRules 
         Height          =   1455
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   2566
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         AllowUserResizing=   1
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   1980
      Width           =   1215
   End
End
Attribute VB_Name = "frmValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998-2001. All Rights Reserved
'   File:       frmValidation.frm
'   Author:     Andrew Newbigging  March 1998
'   Purpose:    Display validation rules for a MACRO question
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'  WillC    Changed the Following where present from Integer to Long  ClinicalTrialId
'           CRFPageId,VisitId,CRFElementID
'   Mo Morris   30/3/00 Changes around the sizing of the grid columns and the overall
'           size of the form and the manner in which it handles re-sizing of the form
'           sub AdjustColumnWidth added
'           sub Form_Resize added
'   TA 17/04/2000:  Adjustments to form and resizing as part of standardisation
' MACRO 2.2
'   NCJ 1 Oct 01 - Updated to cope with 2.2
' MACRO 3.0
'   TA 3/7/02: Dummy sub SetUp removed
'------------------------------------------------------------------------------------'


Option Explicit

'store form height and width
Private mlHeight As Long
Private mlWidth As Long

'---------------------------------------------------------------------
Public Sub Display(oElement As eFormElementRO)
'---------------------------------------------------------------------
' Display the validation rules for this element
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    If PopulateRules(oElement) Then
        'data returned
        'Note that maximum column widths will be calculated by calls
        'to AdjustColumnWidth from within PopulateRules
        
        With grdRules
            If .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + 755 < Screen.Width Then
                'adjust form width to fit (and 375 for any scroll bar)
                Me.Width = .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + 755
            Else
                'size according to rule to be determined
            End If
        End With
        
        'save minimum width and size
        mlHeight = Me.Height
        mlWidth = Me.Width
        
        FormCentre Me
        HourglassSuspend
        Me.Show vbModal
        HourglassResume
    Else
        Call DialogInformation("There are no Validation rules for this question")
    End If
    
Exit Sub
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

End Sub


'---------------------------------------------------------------------
Private Sub cmdHide_Click()
'---------------------------------------------------------------------
' Close form
'---------------------------------------------------------------------
    
    Unload Me
    
End Sub

'---------------------------------------------------------------------
Private Function PopulateRules(oElement As eFormElementRO) As Boolean
'---------------------------------------------------------------------
' Populate grid
' Returns TRUE if there is data to fill grid
' or FALSE if there are no validation rules
' NCJ 1 Oct 01 - Updated to use MACRO 2.2 business objects
'---------------------------------------------------------------------
Dim sItem As String
Dim sMessage As String
Dim bRowsFound As Boolean
Dim oValidation As Validation

    On Error GoTo ErrHandler
    
    HourglassOn
    
    Me.grdRules.Clear
    Me.Caption = "Validation: " & oElement.Name

    If oElement.Validations.Count > 0 Then
        bRowsFound = True
        
        'set up cell behaviour and column headers
        grdRules.ColAlignment(0) = flexAlignLeftCenter
        grdRules.ColAlignment(1) = flexAlignLeftCenter
        grdRules.ColAlignment(2) = flexAlignLeftCenter
        grdRules.TextMatrix(0, 0) = "Type"
        grdRules.TextMatrix(0, 1) = "Validation"
        grdRules.TextMatrix(0, 2) = "Message"
        
        For Each oValidation In oElement.Validations
            
            sMessage = goArezzo.EvaluateExpression(oValidation.MessageExpr)
            If Not goArezzo.ResultOK(sMessage) Then
                ' If we can't evaluate it, just show the message expression itself
                sMessage = oValidation.MessageExpr
            End If
            ' *** Need to convert ValidationType to a string...
            sItem = GetValidationTypeString(oValidation.ValidationType) & _
                    vbTab & _
                    oValidation.ValidationCond & _
                    vbTab & _
                    sMessage
            Me.grdRules.AddItem sItem
            grdRules.Row = grdRules.Rows - 1
            'call AdjustColumnWidth for the newly added data
            AdjustColumnWidth (0)
            AdjustColumnWidth (1)
            AdjustColumnWidth (2)
        Next
    Else
        bRowsFound = False
    End If
    
    Set oValidation = Nothing
    
    HourglassOff
    PopulateRules = bRowsFound

Exit Function
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                "PopulateRules")
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
Private Sub AdjustColumnWidth(ByVal ColumnIndex As Integer)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    grdRules.Col = ColumnIndex
    If grdRules.ColWidth(ColumnIndex) < (TextWidth(Trim(grdRules.Text) & "    ")) Then
        grdRules.ColWidth(ColumnIndex) = (TextWidth(Trim(grdRules.Text) & "    "))
    End If

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "AdjustColumnWidth")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenu.Icon
End Sub

'---------------------------------------------------------------------
Private Sub Form_Resize()
'---------------------------------------------------------------------
' Resize ourselves
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    If Me.Width >= mlWidth Then
        fraRules.Width = Me.ScaleWidth - 120
        grdRules.Width = fraRules.Width - 240
        cmdHide.Left = fraRules.Left + fraRules.Width - cmdHide.Width
    Else
'        Me.Width = mlWidth
    End If
    
    If Me.Height >= mlHeight Then
        fraRules.Height = Me.ScaleHeight - cmdHide.Height - 240
        grdRules.Height = fraRules.Height - 360
        cmdHide.Top = fraRules.Top + fraRules.Height + 120
    Else
'        Me.Height = mlHeight
    End If
    Exit Sub
    
ErrHandler:
    Exit Sub
    
End Sub
