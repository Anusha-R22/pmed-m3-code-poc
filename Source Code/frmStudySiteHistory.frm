VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmStudySiteHistory 
   Caption         =   "Study Site History"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   8115
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   6900
      TabIndex        =   1
      Top             =   3660
      Width           =   1125
   End
   Begin MSComctlLib.ListView lvwHistory 
      Height          =   3435
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   6059
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmStudySiteHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1998-2002. All Rights Reserved
'   File:       frmStudySiteHistory.frm
'   Author:     David Hook, 07/08/2002
'   Purpose:    Show the History for a given study or site
'--------------------------------------------------------------------------------

Option Explicit
Private meDisplay As eDisplayType
Private msStudyName As String
Private mlStudyId As Long
Private msSiteName As String

'--------------------------------------------------------------------------------
Private Sub cmdOK_Click()
'--------------------------------------------------------------------------------
' Close the form
'--------------------------------------------------------------------------------

    Unload Me
    
End Sub

'--------------------------------------------------------------------------------
Private Sub Form_Load()
'--------------------------------------------------------------------------------
' Set up dialog caption
'--------------------------------------------------------------------------------
    ' set icon
    Me.Icon = frmMenu.Icon

    ' Set caption
    Me.Caption = "Study Site History ["
    Select Case meDisplay
        Case eDisplayType.DisplayTrialsBySite
            Me.Caption = Me.Caption & "Site" & " - " & msSiteName & "]"
        Case eDisplayType.DisplaySitesByTrial
            Me.Caption = Me.Caption & "Study" & " - " & msStudyName & "]"
        Case Else
            Call DialogError("Unknown display type", "Study Site History Error")
            Unload Me
    End Select

    ' Set up listview columns
    Call SetListView
    
    ' Populate the listview
    Call PopulateListView

End Sub

'--------------------------------------------------------------------------------
Private Sub SetListView()
'--------------------------------------------------------------------------------
' Set up Listview columns depending on display chosen
'--------------------------------------------------------------------------------
Dim clmHeader As ColumnHeader

    On Error GoTo ErrorHandler
    
    Select Case meDisplay
        Case eDisplayType.DisplayTrialsBySite
            Set clmHeader = lvwHistory.ColumnHeaders.Add(, , "Study", 1440)
        Case eDisplayType.DisplaySitesByTrial
            Set clmHeader = lvwHistory.ColumnHeaders.Add(, , "Site", 1440)
    End Select

    Set clmHeader = lvwHistory.ColumnHeaders.Add(, , "Version", 1440)
    Set clmHeader = lvwHistory.ColumnHeaders.Add(, , "Deploy Date", 2200)
    Set clmHeader = lvwHistory.ColumnHeaders.Add(, , "Received Date", 2200)

Exit Sub
ErrorHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "SetListView")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
  End Select
End Sub

'--------------------------------------------------------------------------------
Private Sub PopulateListView()
'--------------------------------------------------------------------------------
' Populate History ListView from TrialSite table
'--------------------------------------------------------------------------------
    Dim oLI As ListItem
    Dim rsMessage As ADODB.Recordset
    Dim sSQL As String
    Dim lVersion As Long

    On Error GoTo ErrorHandler
    
    ' Set up SQL to get basic version info

    ' Get Messages for trial/site combination
    sSQL = "SELECT MessageTimeStamp, MessageParameters, MessageReceived, MessageReceivedTimeStamp FROM Message " _
        & " WHERE TrialSite = '" & msSiteName & "'" _
        & " AND ClinicalTrialId = " & mlStudyId _
        & " AND MessageType = 8"

    Set rsMessage = MacroADODBConnection.Execute(sSQL, -1, adCmdText)
    Do While Not rsMessage.EOF
        ' Populate Listview
        ' if Version is set (number before parameter name...)
        lVersion = GetStudyVersionFromParameterField(rsMessage("MessageParameters"), msStudyName)
        If lVersion > 0 Then
            ' Study / Site name
            Select Case meDisplay
                Case eDisplayType.DisplayTrialsBySite
                    Set oLI = lvwHistory.ListItems.Add(, , msStudyName)
                Case eDisplayType.DisplaySitesByTrial
                    Set oLI = lvwHistory.ListItems.Add(, , msSiteName)
            End Select
            ' Version
            oLI.SubItems(1) = lVersion
            ' Distribute Date
            oLI.SubItems(2) = Format(rsMessage("MessageTimeStamp"), "yyyy/MM/dd hh:mm")
            ' Received date (if applicable)
            If rsMessage("MessageReceivedTimeStamp") <> 0 Then
                oLI.SubItems(3) = Format(rsMessage("MessageReceivedTimeStamp"), "yyyy/MM/dd hh:mm")
            End If
        End If
        rsMessage.MoveNext
    Loop
    rsMessage.Close
    Set rsMessage = Nothing

Exit Sub
ErrorHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "PopulateListView")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
  End Select
End Sub

'--------------------------------------------------------------------------------
Public Sub InitialiseMe(nDisplay As eDisplayType, sSiteName As String, lStudyId As Long, sStudyName As String)
'--------------------------------------------------------------------------------
' Initialise settings
'--------------------------------------------------------------------------------

    meDisplay = nDisplay
    msStudyName = sStudyName
    mlStudyId = lStudyId
    msSiteName = sSiteName

End Sub
