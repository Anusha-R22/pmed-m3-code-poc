VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStudyStatusHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Study Status History"
   ClientHeight    =   4665
   ClientLeft      =   3600
   ClientTop       =   3390
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   7590
   Begin VB.Frame fraStatusHistory 
      Height          =   4120
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   7480
      Begin MSComctlLib.ListView lvwStatusHistory 
         Height          =   3735
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   6588
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6320
      TabIndex        =   0
      Top             =   4200
      Width           =   1215
   End
End
Attribute VB_Name = "frmStudyStatusHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   File:       FrmStudyStatusHistory.frm
'   Author:     Will Casey Feb 29 2000
'   Purpose:    User can see the history of a given study and its changes in status
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
'Revisions:
'   NCJ 31/5/2000   Show ourselves modally (because calling form is modal)
'   ASH 11/06/2002  Minor change to routine RefreshMe (Bug 2.2.14 no.24)
'   RS  27/06/2003  BUG 1756: Display name of user (that changed studyt status)
'                   as stored in the history table
'--------------------------------------------------------------------------------

Option Explicit
Option Base 0
Option Compare Binary

'--------------------------------------------------------------------------------
Private Sub cmdExit_Click()
'--------------------------------------------------------------------------------
' Unload form
'--------------------------------------------------------------------------------

    Unload Me
    
End Sub

'--------------------------------------------------------------------------------
Private Sub Form_Load()
'--------------------------------------------------------------------------------
' Set up headers for the listview
'--------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Me.Icon = frmMenu.Icon
    
    With lvwStatusHistory
            .ColumnHeaders.Add , , "Name"
            .ColumnHeaders.Add , , "Version", 800
            .ColumnHeaders.Add , , "Status"
            .ColumnHeaders.Add , , "User", 400
            .ColumnHeaders.Add , , "Time"
            .View = lvwReport
            .Sorted = True
    End With
    
    Call FormCentre(Me)
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "Form_Load")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub


'-----------------------------------------------------------
Private Sub lvwStatusHistory_ColumnClick(ByVal ColumnHeader As ColumnHeader)
'-----------------------------------------------------------
' Allow the user to sort by column
'-----------------------------------------------------------

    On Error GoTo ErrHandler

    ' When a ColumnHeader object is clicked, the ListView control is
    ' sorted by the subitems of that column.
    ' Set the SortKey to the Index of the ColumnHeader - 1
    lvwStatusHistory.SortKey = ColumnHeader.Index - 1
    
    ' Reverse the sort order
    If lvwStatusHistory.SortOrder = lvwAscending Then
        lvwStatusHistory.SortOrder = lvwDescending
    Else
        lvwStatusHistory.SortOrder = lvwAscending
    End If
    
    ' Set Sorted to True to sort the list.
    lvwStatusHistory.Sorted = True
        
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "lvwStatusHistory_ColumnClick")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub


'--------------------------------------------------------------------------------
Public Sub RefreshMe(nClinicalTrialID As Long, sSelectedClinicalTrialName As String)
'--------------------------------------------------------------------------------
'Populate the listview
'--------------------------------------------------------------------------------

Dim rsTrialHistory As ADODB.Recordset
Dim sName As String
Dim sVersion As String
Dim nStatusId As eTrialStatus
Dim sStatus As String
Dim sUser As String
'Dim sTime As Date
'ASH 11/06/2002 changed from Date to String
Dim sTime As String
Dim oItem As ListItem
    
    On Error GoTo ErrHandler
    
    'Get the history of a trial and its status'
    Set rsTrialHistory = gdsTrialHistory(nClinicalTrialID)
        
    ' Run through the recordset
    With rsTrialHistory
        Do Until .EOF
            'REM 01/09/03 - if there is no version then put 0
            sVersion = RemoveNull(rsTrialHistory!VersionId)
            If sVersion = "" Then sVersion = "0"
            nStatusId = rsTrialHistory!statusId
            'Mo Morris 30/8/01 Db Audit (UserId to UserName)
            'RS 27/06/2003: BUG 1756: Display name of user as in the history table
            'sUser = rsTrialHistory!StudyDefinitionUserName
            sUser = rsTrialHistory!UserName
            sTime = Format(CDate(rsTrialHistory!StatusChangedTimestamp), "yyyy/mm/dd hh:mm:ss")
        
           'Add a row to the listview
            Set oItem = lvwStatusHistory.ListItems.Add()
            
            ' NCJ 31/5/00 - Use new global GetStudyStatusText
            sStatus = GetStudyStatusText(nStatusId)
    
            'item is the firstitem in the row  and the subitems are the other items in the row
             oItem.Text = sSelectedClinicalTrialName
             oItem.SubItems(1) = sVersion
             oItem.SubItems(2) = sStatus
             oItem.SubItems(3) = sUser
             oItem.SubItems(4) = sTime
            'Go to the next row
            rsTrialHistory.MoveNext
         Loop
    End With

    Set rsTrialHistory = Nothing
    
    ' NCJ 31/5/00 - Must show modally because calling form is now modal
    Me.Show vbModal

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "RefreshMe")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

