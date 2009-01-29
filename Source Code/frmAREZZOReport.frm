VERSION 5.00
Begin VB.Form frmAREZZOReport 
   Caption         =   "AREZZO Terms Report"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8595
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4380
   ScaleWidth      =   8595
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   60
      Width           =   1175
   End
   Begin VB.Frame fraTypes 
      Height          =   2835
      Left            =   60
      TabIndex        =   2
      Top             =   540
      Width           =   2115
      Begin VB.CheckBox chkRegDetails 
         Caption         =   "Registration conditions"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   2340
         Width           =   1935
      End
      Begin VB.CheckBox chkSubjLabels 
         Caption         =   "Subject details"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1980
         Width           =   1635
      End
      Begin VB.CheckBox chkEFormLabels 
         Caption         =   "EForm labels"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1620
         Width           =   1275
      End
      Begin VB.CheckBox chkSkips 
         Caption         =   "Collect if conditions"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1260
         Width           =   1815
      End
      Begin VB.CheckBox chkValMessages 
         Caption         =   "Validation messages"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   900
         Width           =   1755
      End
      Begin VB.CheckBox chkValidations 
         Caption         =   "Validations"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   540
         Width           =   1395
      End
      Begin VB.CheckBox chkDerivations 
         Caption         =   "Derivations"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   180
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   7380
      TabIndex        =   1
      Top             =   3960
      Width           =   1175
   End
   Begin VB.TextBox txtReport 
      Height          =   3795
      Left            =   2280
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   60
      Width           =   6255
   End
End
Attribute VB_Name = "frmAREZZOReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'----------------------------------------------------------------------------------------'
'   File:       frmAREZZOReport.frm
'   Copyright:  InferMed Ltd. 2003. All Rights Reserved
'   Author:     Nicky Johns, June 2003
'   Purpose:    Displays a report on all AREZZO expressions used in the study
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
' REVISIONS:
' NCJ 5-9 June 03 - Initial Development
' NCJ 2 Jul 03 - Corrected Prolog syntax in CheckSemantics
' ic 13/06/2005 added clinical coding
'----------------------------------------------------------------------------------------'

Option Explicit

Private mlTrialID As Long
Private msTrialName As String
Private mbMsgShown As Boolean

' The gap between controls
Private Const mlGAP As Long = 60
Private Const msTAB As String = "   "
    
'----------------------------------------------------------------------------------------'
Public Sub Display(lTrialId As Long, sTrialName As String)
'----------------------------------------------------------------------------------------'
' Display report on all AREZZO terms used in the study
'----------------------------------------------------------------------------------------'

    ' Are we on a different study?
    If lTrialId <> mlTrialID Or txtReport.Text = "" Then
        mlTrialID = lTrialId
        msTrialName = sTrialName
        Call ResetOptions
        Call cmdRefresh_Click
    End If
    
    Me.Show
    Me.ZOrder
    
    Exit Sub
    
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmArezzoReport.Display"

End Sub

'----------------------------------------------------------------------------------------'
Public Sub Clear()
'----------------------------------------------------------------------------------------'
' Clear the text area and reset the options
' (when closing a study)
'----------------------------------------------------------------------------------------'

    txtReport.Text = ""
    Call ResetOptions

End Sub

'----------------------------------------------------------------------------------------'
Private Sub cmdClose_Click()
'----------------------------------------------------------------------------------------'
' Close the window
'----------------------------------------------------------------------------------------'

    Me.Hide

End Sub

'----------------------------------------------------------------------------------------'
Private Sub ReportWarning(sWarn As String, sHeader As String, sTerm As String)
'----------------------------------------------------------------------------------------'
' Add a warning message to the report only if sWarn > ""
' sHeader is first line, sTerm is term which caused the warning
' Sets mbMsgShown = True if sWarn > ""
'----------------------------------------------------------------------------------------'

    If sWarn > "" Then
        mbMsgShown = True
        ' There was something to say
        Logit sHeader
        Logit msTAB & sTerm
        Logit sWarn
        Logit
    End If

End Sub

'----------------------------------------------------------------------------------------'
Private Sub CheckTheWorld(ByVal lTrialId As Long)
'----------------------------------------------------------------------------------------'
' Check every derivation, validation, skip condition and eForm label
' ic 13/06/2005 added clinical coding
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim rsExprs As ADODB.Recordset
Dim sExpr As String
Dim sChecked As String
Dim sCode As String
Dim sFileName As String
Dim sArezzoType As String
Dim sTab As String

    On Error GoTo ErrLabel
    
    txtReport.Text = ""
    Logit msTrialName
    Logit "AREZZO Terms Report " & Format(Now, "dd/mm/yyyy hh:mm:ss") & vbCrLf
    mbMsgShown = False
     
    ' Do the derivations
    If chkDerivations.Value = vbChecked Then
        sSQL = "SELECT DataItemId, DataItemCode, DataType, Derivation FROM DataItem " _
                & " WHERE ClinicalTrialId = " & lTrialId _
                & " AND Derivation IS NOT NULL " _
                & " ORDER BY DataItemCode "
        Set rsExprs = New ADODB.Recordset
        rsExprs.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        Do While Not rsExprs.EOF
            sExpr = RemoveNull(rsExprs!Derivation)
            If sExpr > "" Then
                Select Case rsExprs!DataType
                'ic 13/06/2005 clinical coding: added thesaurus datatype
                Case DataType.Text, DataType.Thesaurus
                    sArezzoType = "string"
                Case DataType.IntegerData, DataType.Real, DataType.LabTest
                    sArezzoType = "numeric"
                Case DataType.Date
                    sArezzoType = "temporal"
                End Select
                Call ReportWarning(CheckSemantics(sExpr, True, sArezzoType), _
                                "Derivation for " & rsExprs!DataItemCode, sExpr)
            End If
            rsExprs.MoveNext
        Loop
        
        rsExprs.Close
    End If
    
    ' Do the validations & messages
    If (chkValidations.Value = vbChecked) Or (chkValMessages.Value = vbChecked) Then
        sSQL = "SELECT DataItem.DataItemId, DataItemCode, DataItemValidation, ValidationMessage " _
                & " FROM DataItem, DataItemValidation " _
                & " WHERE DataItem.ClinicalTrialId = " & lTrialId _
                & " AND DataItem.DataItemId = DataItemValidation.DataItemId " _
                & " AND DataItem.ClinicalTrialId = DataItemValidation.ClinicalTrialId " _
                & " ORDER BY DataItemCode "
        Set rsExprs = New ADODB.Recordset
        rsExprs.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        Do While Not rsExprs.EOF
            sCode = rsExprs!DataItemCode
            If chkValidations.Value = vbChecked Then
                sExpr = RemoveNull(rsExprs!DataItemValidation)
                If sExpr > "" Then
                    Call ReportWarning(CheckSemantics(sExpr, False), _
                                "Validation for " & sCode, sExpr)
                End If
            End If
            If chkValMessages.Value = vbChecked Then
                sExpr = RemoveNull(rsExprs!ValidationMessage)
                If sExpr > "" Then
                    Call ReportWarning(CheckSemantics(sExpr, True), _
                                "Validation message for " & sCode, sExpr)
                End If
            End If
            rsExprs.MoveNext
        Loop
        
        rsExprs.Close
   End If
   
    ' Now the skip conditions
    If chkSkips.Value = vbChecked Then
        sSQL = "SELECT SkipCondition, CRFPageCode FROM CRFElement, CRFPage " _
                & " WHERE CRFElement.ClinicalTrialId = " & lTrialId _
                & " AND CRFElement.CRFPageId = CRFPage.CRFPageId " _
                & " AND CRFElement.ClinicalTrialId = CRFPage.ClinicalTrialId " _
                & " ORDER BY CRFPageCode "
        Set rsExprs = New ADODB.Recordset
        rsExprs.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        Do While Not rsExprs.EOF
            sExpr = RemoveNull(rsExprs!SkipCondition)
            If sExpr > "" Then
                Call ReportWarning(CheckSemantics(sExpr, False), _
                            "Skip condition on eForm " & rsExprs!CRFPageCode, sExpr)
            End If
            rsExprs.MoveNext
        Loop
        
        rsExprs.Close
    End If
    
    ' Now the subject stuff
    If chkSubjLabels.Value = vbChecked Then
        sSQL = "SELECT TrialSubjectLabel, DOBExpr, GenderExpr FROM StudyDefinition " _
                & " WHERE ClinicalTrialId = " & lTrialId
        Set rsExprs = New ADODB.Recordset
        rsExprs.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not rsExprs.EOF Then
        
            sExpr = RemoveNull(rsExprs!TrialSubjectLabel)
            If sExpr > "" Then
                Call ReportWarning(CheckSemantics(sExpr, True), _
                            "Subject label:", sExpr)
            End If
            
            sExpr = RemoveNull(rsExprs!DOBExpr)
            If sExpr > "" Then
                Call ReportWarning(CheckSemantics(sExpr, True), _
                            "Subject DOB Expression:", sExpr)
            End If
            
            sExpr = RemoveNull(rsExprs!GenderExpr)
            If sExpr > "" Then
                Call ReportWarning(CheckSemantics(sExpr, True, "numeric"), _
                            "Subject Gender Expression:", sExpr)
            End If
            
        End If
        
        rsExprs.Close
    
    End If
    
    ' Do the eForm labels
    If chkEFormLabels.Value = vbChecked Then
        sSQL = "SELECT CRFPageLabel, CRFPageCode FROM CRFPage " _
                & " WHERE ClinicalTrialId = " & lTrialId _
                & " ORDER BY CRFPageCode "
        Set rsExprs = New ADODB.Recordset
        rsExprs.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        Do While Not rsExprs.EOF
            sExpr = RemoveNull(rsExprs!CRFPageLabel)
            If sExpr > "" Then
                Call ReportWarning(CheckSemantics(sExpr, True, ""), _
                            "EForm label for " & rsExprs!CRFPageCode, sExpr)
            End If
            rsExprs.MoveNext
        Loop
        
        rsExprs.Close
        
    End If
    
    ' Now the Registration stuff
    If chkRegDetails.Value = vbChecked Then
        ' Uniqueness checks
        sSQL = "SELECT CheckCode, Expression FROM UniquenessCheck " _
                & " WHERE ClinicalTrialId = " & lTrialId
        Set rsExprs = New ADODB.Recordset
        rsExprs.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        Do While Not rsExprs.EOF
        
            sExpr = RemoveNull(rsExprs!Expression)
            If sExpr > "" Then
                Call ReportWarning(CheckSemantics(sExpr, True), _
                            "Uniqueness check - " & RemoveNull(rsExprs!CheckCode), sExpr)
            End If
                  
            rsExprs.MoveNext
        Loop
        
        rsExprs.Close
    
        ' Eligibility conditions
        sSQL = "SELECT EligibilityCode, Condition FROM Eligibility " _
                & " WHERE ClinicalTrialId = " & lTrialId
        Set rsExprs = New ADODB.Recordset
        rsExprs.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        Do While Not rsExprs.EOF
        
            sExpr = RemoveNull(rsExprs!Condition)
            If sExpr > "" Then
                Call ReportWarning(CheckSemantics(sExpr, False), _
                            "Eligibility condition - " & RemoveNull(rsExprs!EligibilityCode), sExpr)
            End If
                  
            rsExprs.MoveNext
        Loop
        
        rsExprs.Close
    
        ' Suffix and prefix
        sSQL = "SELECT Prefix, Suffix FROM SubjectNumbering " _
                & " WHERE ClinicalTrialId = " & lTrialId
        Set rsExprs = New ADODB.Recordset
        rsExprs.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not rsExprs.EOF Then
        
            sExpr = RemoveNull(rsExprs!Prefix)
            If sExpr > "" Then
                Call ReportWarning(CheckSemantics(sExpr, True), _
                            "Registration prefix", sExpr)
            End If
                  
            sExpr = RemoveNull(rsExprs!Suffix)
            If sExpr > "" Then
                Call ReportWarning(CheckSemantics(sExpr, True), _
                            "Registration suffix", sExpr)
            End If
            
        End If
        
        rsExprs.Close
    
    End If
 
    Set rsExprs = Nothing
    
    If Not mbMsgShown Then
        ' No messages shown
        Logit "No errors or warnings were found"
        Logit
    End If
    
    Logit Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    Exit Sub
    
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmArezzoReport.CheckTheWorld"

End Sub

'---------------------------------------------------------------------
Private Function CheckSemantics(sTerm As String, bExpr As Boolean, Optional sType As String = "") As String
'---------------------------------------------------------------------
' Check the semantics of an AREZZO term
' Pass bExpr = TRUE for expression, or FALSE for condition
' NCJ 2 Jul 03 - Corrected Prolog syntax!
'---------------------------------------------------------------------
Dim sQuery As String
Dim sResultCode As String
Dim sCheckText As String

    On Error GoTo ErrLabel
    
    If Not bExpr Then
        ' Condition
        sQuery = "pdl_check_cond( `" & sTerm & "` ). "
        sCheckText = goALM.GetPrologResult(sQuery, sResultCode)
    Else
        ' Expression
        sQuery = "pdl_check_expr( `" & sTerm & "`"
        If sType <> "" Then
            sQuery = sQuery & ", " & sType
        End If
        sQuery = sQuery & " ). "
        
        sCheckText = goALM.GetPrologResult(sQuery, sResultCode)
    
    End If
    CheckSemantics = sCheckText

Exit Function

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmArezzoReport.CheckSemantics"

End Function

'---------------------------------------------------------------------
Private Sub Logit(Optional sText As String = "")
'---------------------------------------------------------------------
' Add text to the Report text box, followed by linefeed
'---------------------------------------------------------------------

    txtReport.Text = txtReport.Text & sText & vbCrLf

End Sub

'---------------------------------------------------------------------
Private Sub cmdRefresh_Click()
'---------------------------------------------------------------------
' Refresh the report
'---------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    If IsSomethingChecked Then
        cmdRefresh.Enabled = False
        HourglassOn
        Call CheckTheWorld(mlTrialID)
        HourglassOff
        cmdRefresh.Enabled = True
    End If
    
Exit Sub

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmArezzoReport.cmdRefresh_Click"

End Sub

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------

    ' Force its size otherwise for some unexplained reason
    ' it chooses a different size for itself!
    Me.Height = 4815
    Me.Width = 8730
    
    Call FormCentre(Me, frmMenu)
    Set Me.Icon = frmMenu.Icon
    
    ' Initialise the options
    Call ResetOptions
    
End Sub

'---------------------------------------------------------------------
Private Sub ResetOptions()
'---------------------------------------------------------------------

    ' Initialise the options
    chkDerivations.Value = vbChecked
    chkValidations.Value = vbChecked
    chkValMessages.Value = vbChecked
    chkSkips.Value = vbChecked
    chkEFormLabels.Value = vbChecked
    chkSubjLabels.Value = vbChecked
    chkRegDetails.Value = vbChecked
    
End Sub

'---------------------------------------------------------------------
Private Function IsSomethingChecked() As Boolean
'---------------------------------------------------------------------
' Have they selected something?
'---------------------------------------------------------------------

    IsSomethingChecked = True
    
    If chkDerivations.Value = vbChecked Then Exit Function
    If chkValidations.Value = vbChecked Then Exit Function
    If chkValMessages.Value = vbChecked Then Exit Function
    If chkSkips.Value = vbChecked Then Exit Function
    If chkEFormLabels.Value = vbChecked Then Exit Function
    If chkSubjLabels.Value = vbChecked Then Exit Function
    If chkRegDetails.Value = vbChecked Then Exit Function

    IsSomethingChecked = False
    DialogWarning "Please select at least one term type"
    
End Function

'---------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'---------------------------------------------------------------------

    ' If they clicked the Close box, just hide the window rather than unload it
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        Me.Hide
    End If

End Sub

'---------------------------------------------------------------------
Private Sub Form_Resize()
'---------------------------------------------------------------------

    On Error Resume Next
    
        'check that form has not been minimised
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Call FitToWidth(Me.ScaleWidth)

    Call FitToHeight(Me.ScaleHeight)


End Sub

'--------------------------------------------------------------------
Private Sub FitToWidth(ByVal lWinWidth As Long)
'--------------------------------------------------------------------
' Fit the controls into the given window width
' Assume the width is not below the minimum
'--------------------------------------------------------------------

    ' Size the message box
    txtReport.Width = Max(lWinWidth - txtReport.Left - mlGAP, _
                            cmdClose.Width)
    ' Move the Close button
    cmdClose.Left = txtReport.Left + txtReport.Width - cmdClose.Width

End Sub

'--------------------------------------------------------------------
Private Sub FitToHeight(ByVal lWinHeight As Long)
'--------------------------------------------------------------------
' Fit the controls into the given window width
' Assume the height is not below the minimum
'--------------------------------------------------------------------

    ' Move the Close button
    cmdClose.Top = Max(fraTypes.Top + fraTypes.Height + mlGAP, _
                        lWinHeight - cmdClose.Height - mlGAP)
    ' Size the message box
    txtReport.Height = cmdClose.Top - txtReport.Top - mlGAP

End Sub

