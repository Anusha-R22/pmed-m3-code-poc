VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDataCommunication 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data Communication"
   ClientHeight    =   4275
   ClientLeft      =   4890
   ClientTop       =   4035
   ClientWidth     =   5760
   Icon            =   "frmDataCommunication.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   3360
      Width           =   5295
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Close"
         Height          =   375
         Left            =   3840
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3015
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5318
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "By site"
      TabPicture(0)   =   "frmDataCommunication.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cboSites"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtLastMessage"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "By period"
      TabPicture(1)   =   "frmDataCommunication.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(2)=   "Label5"
      Tab(1).Control(3)=   "txtDays"
      Tab(1).Control(4)=   "lstSites"
      Tab(1).Control(5)=   "Frame2"
      Tab(1).ControlCount=   6
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   -74640
         TabIndex        =   11
         Top             =   2040
         Width           =   4935
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Height          =   375
            Left            =   3480
            TabIndex        =   12
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.ListBox lstSites 
         Height          =   1035
         Left            =   -73560
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox txtDays 
         Height          =   285
         Left            =   -71520
         TabIndex        =   0
         Top             =   570
         Width           =   615
      End
      Begin VB.TextBox txtLastMessage 
         Enabled         =   0   'False
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   2040
         Width           =   3015
      End
      Begin VB.ComboBox cboSites 
         Height          =   315
         Left            =   360
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Last contact:"
         Height          =   375
         Left            =   -74520
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Choose a remote site to find out the last time the site communicated data to this centre."
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label Label3 
         Caption         =   "days."
         Height          =   255
         Left            =   -70800
         TabIndex        =   8
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Sites that have not made contact in the last"
         Height          =   255
         Left            =   -74640
         TabIndex        =   5
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Date and time of last contact:"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1800
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmDataCommunication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999 - 2000. All Rights Reserved
'   File:       frmDataCommunication.frm
'   Author:     Will Casey, July 1999
'   Purpose:    To enable the administrator to find out when a site last made contact
'               with the central site by site or for a given period of days.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
' Revisions:
'   Nicky & Mo, 16 Jun 00 - Corrected cmdFind_Click code
'   WillC   SR3561 20/6/00 fixed rte when no entry in days box and user clicks find
'           Also changed the error handlers to use ExitMacro and MacroEnd
'   DPH 03/05/2002 Get site transfer details from LogDetails and Data Integrity info
'------------------------------------------------------------------------------------'
Option Explicit

Private msSite As String


'---------------------------------------------------------------------
Private Sub cmdExit_Click()
'---------------------------------------------------------------------

'---------------------------------------------------------------------

    Unload Me
    
End Sub

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------
' Get the list of sites
'---------------------------------------------------------------------
    On Error GoTo ErrHandler

    Call RefreshSites
    
    'WillC   20/6/00
    cmdFind.Enabled = False
    
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "Form_Load")
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
Private Sub cboSites_Click()
'---------------------------------------------------------------------
' Clear the previous entry and place the chosen site in the variable
'---------------------------------------------------------------------
    On Error GoTo ErrHandler

    txtLastMessage.Text = ""
    msSite = cboSites.Text
    
    Call RefreshList(msSite)
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cboSites_Click")
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
Private Sub RefreshList(msSite As String)
'---------------------------------------------------------------------
' Find the time of last incoming message from a given site. MessageIn is
' an enumerator 1 is message in.Add the time to the textbox.
'---------------------------------------------------------------------
' REVISIONS
' DPH 03/05/2002 Get site transfer details from LogDetails and Data Integrity info
'---------------------------------------------------------------------

Dim sSQL As String
Dim rsMessageTime  As ADODB.Recordset
    
    On Error GoTo ErrHandler

'    sSQL = "SELECT Distinct MessageTimestamp FROM Message "
'    sSQL = sSQL & "WHERE TrialSite = '" & msSite & "'"
'    sSQL = sSQL & "AND MessageDirection = " & MessageIn '1 is message in
    sSQL = "SELECT DISTINCT LogDateTime FROM LogDetails WHERE TaskId = 'DataIntegrity' " _
        & "AND " 'LIKE '"
    sSQL = sSQL & GetSQLStringLike("LogMessage", "Site " & msSite)
    sSQL = sSQL & " ORDER BY LogDateTime DESC"
    
    Set rsMessageTime = New ADODB.Recordset
    rsMessageTime.Open sSQL, MacroADODBConnection, adOpenForwardOnly, , adCmdText
    
'    With rsMessageTime
'        Do Until .EOF = True
'            If Not IsNull(.Fields("MessageTimeStamp")) Then
'                txtLastMessage.Text = Format(CDate(.Fields("MessageTimeStamp")), "yyyy/mm/dd hh:mm:ss")
'            End If
'            .MoveNext
'        Loop
'    End With

    ' DPH 03/05/2002 - Get first record
    If Not rsMessageTime.EOF Then
        txtLastMessage.Text = Format(CDate(rsMessageTime.Fields("LogDateTime")), "yyyy/mm/dd hh:mm:ss")
    Else
        txtLastMessage.Text = "Site " & msSite & " has never made contact"
    End If
    rsMessageTime.Close
    Set rsMessageTime = Nothing
     
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "RefreshList")
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
Private Sub RefreshSites()
'---------------------------------------------------------------------
' Add all the Sites to the combobox
'---------------------------------------------------------------------

Dim sSQL As String
Dim rsTrialSites  As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    sSQL = "SELECT DISTINCT TrialSite.TrialSite FROM TrialSite, Site" _
         & " WHERE TrialSite.TrialSite = Site.Site " _
         & " AND Site.SiteLocation = " & SiteLocation.ESiteRemote
    Set rsTrialSites = New ADODB.Recordset
    
    rsTrialSites.Open sSQL, MacroADODBConnection, adOpenForwardOnly, , adCmdText
    
    With rsTrialSites
        Do Until .EOF = True
         cboSites.AddItem .Fields("Trialsite")
                .MoveNext
        Loop
    End With
    
    Set rsTrialSites = Nothing
        
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "RefreshSites")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub



'------------------------------------------------------------------------------'
Private Sub txtDays_Change()
'------------------------------------------------------------------------------'
' WillC 20/6/00 if the user enters a non numeric entry in the Days box then leave the find button
' disabled...
'------------------------------------------------------------------------------'
Dim nDays As Integer
Dim sDays As String
'Dim X As String
'Dim sNumbers As String
    On Error GoTo ErrHandler
    
    cmdFind.Enabled = False
    sDays = txtDays.Text
       
    'WillC SR3561  20/6/00
    If Not sDays = vbNullString And IsNumeric(sDays) Then
        nDays = Val(CInt(sDays))    'check for overflow errors
        If Val(nDays) = Val(sDays) Then
           cmdFind.Enabled = True
        End If
    Else
        cmdFind.Enabled = False
    End If
        
'  WillC 20/6/00  Use IsNumeric as above instead
'    X = Chr(KeyAscii)
'    sNumbers = "0123456789" & vbBack & vbCr
'
'    If InStr(sNumbers, X) = 0 Then
'        MsgBox "Only Numeric values are allowed", vbOKOnly, "MACRO"
'        KeyAscii = 8 ' backspace clears the Offending character
'    End If
'
'    If KeyAscii = Asc(vbCr) Then
'        cmdFind.Value = True
'    End If

Exit Sub
ErrHandler:
    Select Case Err.Number
        Case 6  ' Trap any overflow errors
            MsgBox "The number entered is too large. Please enter a smaller number.", vbInformation, "MACRO"
            txtDays.Text = vbNullString
            Exit Sub
        Case Else
            Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtDays_Change")
                  Case OnErrorAction.Ignore
                      Resume Next
                  Case OnErrorAction.Retry
                      Resume
                  Case OnErrorAction.QuitMACRO
                      Call ExitMACRO
                      Call MACROEnd
             End Select
        End Select
        
End Sub

'------------------------------------------------------------------------------'
Private Sub cmdFind_Click()
'------------------------------------------------------------------------------'
' Find which sites havent communicated in the last number of days as specified.
' We cast the date into a long to stop the database from changing the format from
' dd/mm/yyyy to mm/dd/yyyy when we do a comparison on dates which it will do if
' it can this seems to be an access bug ie it will change 12/11/1999(Nov 12)
'  to 11/12/1999 (Dec 11)
'
'Re-written by Mo Morris so that it displays sites for which there has been no communication
'------------------------------------------------------------------------------'
' REVISIONS
' DPH 03/05/2002 Get site transfer details from LogDetails and Data Integrity info
'------------------------------------------------------------------------------'
Dim sDaysSQLString As String
Dim sSQL As String
Dim rsMessageDates As ADODB.Recordset
Dim rsSites As ADODB.Recordset
'Dim n As Integer
'Dim dtAfter As Variant

    On Error GoTo ErrHandler

    sDaysSQLString = ConvertLocalNumToStandard(CStr(CDbl(Now) - CLng(txtDays.Text)))
    
    lstSites.Clear
    
    sSQL = "SELECT Site FROM Site" _
        & " WHERE Site.SiteLocation = " & SiteLocation.ESiteRemote _
        & " ORDER BY site"
    Set rsSites = New ADODB.Recordset
    rsSites.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Do Until rsSites.EOF
'        sSQL = "SELECT MessageTimeStamp FROM Message" _
'            & " WHERE TrialSite = '" & rsSites!site & "'" _
'            & " AND MessageTimeStamp > " & sDaysSQLString _
'            & " AND MessageType = " & ExchangeMessageType.PatientData

        sSQL = "SELECT LogDateTime FROM LogDetails WHERE TaskId = 'DataIntegrity' AND " ' LogMessage LIKE '"
        sSQL = sSQL & GetSQLStringLike("LogMessage", "Site " & rsSites!Site)
        sSQL = sSQL & " AND LogDateTime > " & sDaysSQLString
        
        Set rsMessageDates = New ADODB.Recordset
        rsMessageDates.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
        If rsMessageDates.RecordCount = 0 Then
            lstSites.AddItem rsSites!Site
        End If
        rsMessageDates.Close
        Set rsMessageDates = Nothing
        rsSites.MoveNext
    Loop
    rsSites.Close
    Set rsSites = Nothing
    

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdFind_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

    
    
    

'    nDays = -Val(txtDays.Text)
'    dtAfter = DateAdd("d", nDays, Now)
'    sSQL = "SELECT Trialsite, MessageTimeStamp FROM Message "
'    ' WillC 4/2/00 Changed to cope with local settings problem due to commas
'    sSQL = sSQL & " WHERE MessageTimestamp > " & ConvertLocalNumToStandard(CStr(CDbl(dtAfter)))
'    sSQL = sSQL & " AND MessageDirection = 1"
'    Set rsMessageDates = New ADODB.Recordset
'    rsMessageDates.Open sSQL, MacroADODBConnection, adOpenForwardOnly, , adCmdText
'    With rsMessageDates
'    n = rsMessageDates.RecordCount
'       Do Until .EOF = True
'            lstSites.AddItem (.Fields("TrialSite") & "  at    " & Format(.Fields("MessageTimeStamp"), "hh:MM dd/mm/yyyy"))
'            .MoveNext
'        Loop
'    End With
'    Set rsMessageDates = Nothing

End Sub

