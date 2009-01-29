VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAuditTrail 
   Caption         =   "Audit Trail"
   ClientHeight    =   4470
   ClientLeft      =   5550
   ClientTop       =   6315
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6630
   Begin VB.Frame fraAudit 
      Caption         =   " Audit Trail "
      Height          =   3855
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   6495
      Begin MSFlexGridLib.MSFlexGrid grdAuditTrail 
         Height          =   3495
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   6165
         _Version        =   393216
         Rows            =   1
         Cols            =   6
         FixedCols       =   0
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         AllowUserResizing=   3
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5340
      TabIndex        =   0
      Top             =   4020
      Width           =   1215
   End
   Begin VB.Label lblCommentSize 
      AutoSize        =   -1  'True
      Caption         =   "CommentSize"
      Height          =   195
      Left            =   2820
      TabIndex        =   1
      Top             =   4080
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "frmAuditTrail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999-2000. All Rights Reserved
'   File:       frmAuditTrial.frm
'   Author:     Steve Morris, November 1999
'   Purpose:    To show an audit trail for a data item during data entry.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'Revisions:
'   SDM 29/11/99    Created for SR1494
'   NCJ 30 Nov 99 Ensure timestamps are cast as dates
'   WillC 11 / 12 / 99
'       Changed the Following where present from Integer to Long  ClinicalTrialId
'       CRFPageId,VisitId,CRFElementID
'   Mo Morris   3/3/00  SR3064
'       reasons for change added to the forms listview
'       initial size of form changed
'   NCJ 10 Mar 00 - Make sure comment field is correct height
'   TA 17/04/2000:  Adjustments to form and resizing as part of standardisation
'   TA 28/04/2000:  Column widths now include column headers
'   TA 10/05/2000 SR3416: Lock Status now shown
'   NCJ 1/9/00 - SR 3873 in PopulateAudit
'   TA 19/08/2000:  Normal Range and CTC grade now shown
'   NCJ 26/9/00 - NR/CTC Column corrected
'   NCJ 6/10/00 - Moved NR/CTC in with status; added WarningMessage column;
'                   resize row heights for Comments and WarningMessages
'   NCJ 24/11/00 - Show lab code for LabTest questions
'
' MACRO 2.2
'   NCJ 26 Sep 01 - Updated to use eFormElementRO and eFormInstance
'
' MACRO 3.0
'   NCJ 15 Nov 01 - Include element's RepeatNumber
'   NCJ 2 Jan 02 - Show Repeat Number in window caption
'
' NCJ 9 Jul 2002 - Set form icon - correctly!
' RS 10/10/2002 -   Added display of timezone in timestamp column
'----------------------------------------------------------------------------------------'

Option Explicit

' The columns

Private Const mnCOL_RESPONSE = 0
Private Const mnCOL_STATUS = 1
Private Const mnCOL_LABCODE = 2
Private Const mnCOL_USER = 3
Private Const mnCOL_TIMESTAMP = 4
Private Const mnCOL_WARNMSG = 5
Private Const mnCOL_OVERRULE = 6
Private Const mnCOL_RFC = 7
Private Const mnCOL_COMMENT = 8
Private Const mnCOL_LOCKSTATUS = 9

Private Const mnCOL_LIMIT = 9

' Initial form height and width
Private mlHeight As Long
Private mlWidth As Long

' NCJ 24/11/00 - Store the type of the question
Private mnDataType As Integer

'---------------------------------------------------------------------
Public Sub Setup(oElement As eFormElementRO, nRow As Integer, _
                 oEFI As EFormInstance)
'---------------------------------------------------------------------
' Setup the form
' NCJ 26 Sep 01 - Changed input parameters to take eFormInstance
' NCJ 15 Nov 01 - Added nRow to input parameters
'---------------------------------------------------------------------
Dim lWidth As Long
Dim lCol As Long

    If Not oElement.ElementID > 0 Then Exit Sub
                
    If PopulateAudit(oElement, nRow, _
                    oEFI.VisitInstance.Subject.StudyId, _
                    oEFI.VisitInstance.Subject.Site, _
                    oEFI.VisitInstance.Subject.PersonId, _
                    oEFI.EFormTaskId) Then
        With grdAuditTrail
            For lCol = 0 To mnCOL_LIMIT
                lWidth = lWidth + .ColWidth(lCol)
            Next
            lWidth = lWidth
            If lWidth + 755 < Screen.Width Then
                'adjust form width to fit (add 775 for any scroll bar)
                Me.Width = lWidth + 755
            End If
        End With
        
        'TA: minimum width and height for resizing are set here
        '       values are now 2000 rather than initial display values
        mlHeight = 2000 'Me.Height
        mlWidth = 2000 'Me.Width
        
        FormCentre Me
        HourglassSuspend
        Me.Show vbModal
        HourglassResume
    Else
        Call DialogInformation("There is no Audit Trail for this question")
    End If

End Sub

'---------------------------------------------------------------------
Private Sub cmdHide_Click()
'---------------------------------------------------------------------

    Unload Me

End Sub

'---------------------------------------------------------------------
Private Function PopulateAudit(oElement As eFormElementRO, _
                                ByVal nRptNo As Integer, _
                                ByVal lClinicalTrialId As Long, _
                                ByVal TrialSite As String, _
                                ByVal PersonId As Integer, _
                                ByVal CRFPageTaskId As Long) As Boolean
'---------------------------------------------------------------------
'   Get data from database and populate grid
'   Output:
'           function - true if data returned
' NCJ 2 May 00 SR 3399 - Changed CRFPageId to CRFPageTaskId
' NCJ 1/9/00 - SR 3873 Check user code found in Security.mdb
' NCJ 24/11/00 - Show lab code for LabTest questions
' NCJ 26/9/01 - Pass oElement instead of nElementID
' NCJ 15/11/01 - Pass nRptNo as well
' NCJ 2/1/02 - Show Repeat Number in caption
'---------------------------------------------------------------------
Dim sSQL As String
Dim sItem As String
Dim rsWarning As ADODB.Recordset
Dim rsUserName As ADODB.Recordset
Dim bRowsFound As Boolean
Dim sUsername As String
Dim sStatus As String
Dim sNRStatus As String
Dim sWarnMsg As String
Dim sComments As String
Dim lRowHeight As Long
Dim sTimestamp As String    ' RS 10/10/2002: Pre-format Timestring
Dim oTimezone As Timezone

    Set oTimezone = New Timezone
        
    Call HourglassOn
    
    Me.Caption = "Audit Trail: " & oElement.Name & " [" & nRptNo & "]"
    ' NCJ 24/11/00
    mnDataType = oElement.DataType
    
    'Mo Morris 3/3/00, ReasonForChange added to SQL statement
    ' NCJ 2/5/00 SR 3399 CRFPageId -> CRFPageTaskID
    ' NCJ 24/11/00 - Added LaboratoryCode
    ' NCJ 15/11/01 - Added RepeatNumber
    ' RS 10/10/2002 - Added Timezone
    sSQL = "SELECT " & _
            "DataItemResponseHistory.ResponseValue, " & _
            "DataItemResponseHistory.ResponseStatus, " & _
            "DataItemResponseHistory.UserName, " & _
            "DataItemResponseHistory.ResponseTimestamp, " & _
            "DataItemResponseHistory.Comments, " & _
            "DataItemResponseHistory.ReasonForChange, " & _
            "DataItemResponseHistory.ValidationMessage, " & _
            "DataItemResponseHistory.OverruleReason, " & _
            "DataItemResponseHistory.LockStatus, " & _
            "DataItemResponseHistory.LabResult, " & _
            "DataItemResponseHistory.CTCGrade, " & _
            "DataItemResponseHistory.LaboratoryCode, " & _
            "DataItemResponseHistory.ResponseTimestamp_TZ " & _
            "FROM DataItemResponseHistory " & _
            "WHERE DataItemResponseHistory.ClinicalTrialId = " & lClinicalTrialId & " " & _
            "AND DataItemResponseHistory.TrialSite = '" & TrialSite & "' " & _
            "AND DataItemResponseHistory.PersonId = " & PersonId & _
            "AND DataItemResponseHistory.CRFElementId = " & oElement.ElementID & " " & _
            "AND DataItemResponseHistory.CRFPageTaskId = " & CRFPageTaskId & " " & _
            "AND DataItemResponseHistory.RepeatNumber = " & nRptNo & " " & _
            "ORDER BY DataItemResponseHistory.ResponseTimestamp DESC"
    Set rsWarning = New ADODB.Recordset
    rsWarning.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rsWarning.EOF Then
        bRowsFound = True
        
        grdAuditTrail.Cols = mnCOL_LIMIT + 1
        grdAuditTrail.Clear
        
        'set up column headers
        grdAuditTrail.TextMatrix(0, mnCOL_RESPONSE) = "Response"
        grdAuditTrail.TextMatrix(0, mnCOL_STATUS) = "Status"
        grdAuditTrail.TextMatrix(0, mnCOL_LABCODE) = "Laboratory"
        grdAuditTrail.TextMatrix(0, mnCOL_USER) = "User Name"
        grdAuditTrail.TextMatrix(0, mnCOL_TIMESTAMP) = "Timestamp"
        grdAuditTrail.TextMatrix(0, mnCOL_WARNMSG) = "Warning"
        grdAuditTrail.TextMatrix(0, mnCOL_OVERRULE) = "Overrule Reason"
        grdAuditTrail.TextMatrix(0, mnCOL_RFC) = "Reasons for change"
        grdAuditTrail.TextMatrix(0, mnCOL_COMMENT) = "Comments"
        grdAuditTrail.TextMatrix(0, mnCOL_LOCKSTATUS) = "Locked/Frozen"
        
        grdAuditTrail.Row = 0
        
        'TA 28/04/2000:   adjust column widths for header
        Call AdjustColumnWidth
       
        rsWarning.MoveFirst
        Do While Not rsWarning.EOF
            
            ' Pick up the user's name (if we can)
            sSQL = "SELECT UserNameFull FROM MACROUser " & _
                    " WHERE UserName = '" & rsWarning.Fields("UserName").Value & "'"
            Set rsUserName = New ADODB.Recordset
            rsUserName.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
            ' NCJ 1/9/00 - SR 3873 Check that user was found
            If Not rsUserName.EOF Then
                ' We found this user in the Security database
                sUsername = rsUserName.Fields("UserNameFull").Value
            Else
                ' User not found so just display the code
                sUsername = rsWarning.Fields("UserName").Value
            End If
            Set rsUserName = Nothing
            
            ' NCJ 6/10/00 - Add NR Status to ordinary status
            sStatus = GetStatusText(Val(rsWarning.Fields("ResponseStatus").Value))
            sNRStatus = GetNRCTCText((rsWarning.Fields("LabResult").Value), (rsWarning.Fields("CTCGrade").Value))
            If sNRStatus > "" Then
                sStatus = sStatus & " [" & sNRStatus & "]"
            End If

            ' RS 10/10/2002: Format timestamp according to user settings
            If GetMACROSetting("timestampdisplay", "storedvalue") = "storedvalue" Then
                sTimestamp = Format(CDate(rsWarning.Fields("ResponseTimestamp").Value), "yyyy/mm/dd hh:mm:ss")
                sTimestamp = sTimestamp & " (GMT" & IIf(rsWarning.Fields("ResponseTimestamp_TZ").Value < 0, "+", "") & _
                                -rsWarning.Fields("ResponseTimestamp_TZ").Value \ 60 & ":" & _
                                Format(Abs(rsWarning.Fields("ResponseTimestamp_TZ").Value) Mod 60, "00") & ")"
            Else
                sTimestamp = Format(CDate(oTimezone.ConvertDateTimeToLocal(rsWarning.Fields("ResponseTimestamp").Value, rsWarning.Fields("ResponseTimestamp_TZ").Value)), "yyyy/mm/dd hh:mm:ss")
            End If


            sWarnMsg = RemoveNull(rsWarning.Fields("ValidationMessage").Value)
            sComments = RemoveNull(rsWarning.Fields("Comments").Value)
            sItem = rsWarning.Fields("ResponseValue").Value & vbTab & _
                    sStatus & vbTab & _
                    RemoveNull(rsWarning.Fields("LaboratoryCode").Value) & vbTab & _
                    sUsername & vbTab & _
                    sTimestamp & vbTab & _
                    sWarnMsg & vbTab & _
                    RemoveNull(rsWarning.Fields("OverruleReason").Value) & vbTab & _
                    RemoveNull(rsWarning.Fields("ReasonForChange").Value) & vbTab & _
                    sComments & vbTab & _
                    GetLockStatusText(rsWarning.Fields("LockStatus").Value)
                    
            grdAuditTrail.AddItem sItem
            
            grdAuditTrail.Row = grdAuditTrail.Rows - 1
            
            ' adjust column widths (including dealing with Laboratory column)
            Call AdjustColumnWidth
            
            ' Adjust row height for Comments & Warning Message
            ' (These can contain carriage returns)
            lRowHeight = grdAuditTrail.RowHeight(grdAuditTrail.Row)
            If Len(sComments) > 0 Then
                ' NCJ 10/3/00 - Set width of lblCommentSize first
                lblCommentSize.Width = grdAuditTrail.ColWidth(mnCOL_COMMENT)
                lblCommentSize.Caption = sComments
                If lRowHeight < lblCommentSize.Height Then
                    lRowHeight = lblCommentSize.Height
                End If
'                grdAuditTrail.RowHeight(grdAuditTrail.Row) = lblCommentSize.Height
            End If
            If Len(sWarnMsg) > 0 Then
                ' NCJ 10/3/00 - Set width of lblCommentSize first
                lblCommentSize.Width = grdAuditTrail.ColWidth(mnCOL_WARNMSG)
                lblCommentSize.Caption = sWarnMsg
                If lRowHeight < lblCommentSize.Height Then
                    lRowHeight = lblCommentSize.Height
                End If
            End If
            If lRowHeight > grdAuditTrail.RowHeight(grdAuditTrail.Row) Then
                ' Set height with fudge factor
                grdAuditTrail.RowHeight(grdAuditTrail.Row) = lRowHeight + (6 * Screen.TwipsPerPixelY)
            End If
            rsWarning.MoveNext
            
        Loop
        
        grdAuditTrail.ColAlignment(mnCOL_COMMENT) = flexAlignCenterTop
        grdAuditTrail.ColAlignment(mnCOL_RFC) = flexAlignCenterTop

         
    Else
        bRowsFound = False
    End If
    Set rsWarning = Nothing
    
    HourglassOff
    
    PopulateAudit = bRowsFound
    
Exit Function
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "PopulateAudit")
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
Private Sub AdjustColumnWidth()
'---------------------------------------------------------------------
'   SDM 06/12/99
'   TA 28/04/2000: now done in loop
' NCJ 24/11/00 - Set Laboratory column width to 0 for non-LabTest qus.
'---------------------------------------------------------------------
Dim nCol As Integer
    
    On Error GoTo ErrHandler
    
    With grdAuditTrail
        For nCol = 0 To mnCOL_LIMIT
            .Col = nCol
            'TA 28/04/2000: width padding now done by adding 12 pixels rather than by appending spaces
            If .ColWidth(nCol) < (TextWidth(Trim(.Text)) + 12 * Screen.TwipsPerPixelX) Then
                .ColWidth(nCol) = (TextWidth(Trim(.Text)) + 12 * Screen.TwipsPerPixelX)
            End If
            .ColAlignment(nCol) = flexAlignLeftCenter
        Next
        ' Deal with LabCode column
        If mnDataType <> eDataType.LabTest Then
            .ColWidth(mnCOL_LABCODE) = 0
        End If
    End With

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "AdjustColumnWidth")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
End Sub

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------
' NCJ 9 Jul 2002 - Set form icon - correctly!
'---------------------------------------------------------------------

    Me.Icon = frmMenu.Icon

End Sub

'---------------------------------------------------------------------
Private Sub Form_Resize()
'---------------------------------------------------------------------
'Mo Morris 6/3/00   Minimum size set to 6900 * 5000
'---------------------------------------------------------------------

On Error GoTo ErrHandler

    If Me.Width >= mlWidth Then
        fraAudit.Width = Me.ScaleWidth - 120
        grdAuditTrail.Width = fraAudit.Width - 240
        cmdHide.Left = fraAudit.Left + fraAudit.Width - cmdHide.Width
    Else
'        Me.Width = mlWidth
    End If
    
    If Me.Height >= mlHeight Then
        fraAudit.Height = Me.ScaleHeight - cmdHide.Height - 240
        grdAuditTrail.Height = fraAudit.Height - 360
        cmdHide.Top = fraAudit.Top + fraAudit.Height + 120
    Else
'        Me.Height = mlHeight
    End If
    Exit Sub
    
ErrHandler:
    Exit Sub
    
End Sub


