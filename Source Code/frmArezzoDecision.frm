VERSION 5.00
Begin VB.Form frmArezzoDecision 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Decision"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9810
   ControlBox      =   0   'False
   Icon            =   "frmArezzoDecision.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   300
      TabIndex        =   11
      Top             =   4440
      Width           =   1400
   End
   Begin VB.CommandButton cmdHideExplain 
      Caption         =   "Hide Explanation"
      Height          =   495
      Left            =   7800
      TabIndex        =   10
      Top             =   4440
      Width           =   1400
   End
   Begin VB.TextBox txtExplain 
      Height          =   3735
      Left            =   5160
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Text            =   "frmArezzoDecision.frx":030A
      Top             =   480
      Width           =   4335
   End
   Begin VB.Frame fraCand 
      Caption         =   "Frame1"
      Height          =   975
      Left            =   240
      TabIndex        =   5
      Top             =   3200
      Width           =   4575
      Begin VB.TextBox txtCandInfo 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   170
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "frmArezzoDecision.frx":0310
         Top             =   240
         Width           =   4270
      End
   End
   Begin VB.CommandButton cmdExplain 
      Caption         =   "Explain"
      Height          =   495
      Left            =   1860
      TabIndex        =   4
      Top             =   4440
      Width           =   1400
   End
   Begin VB.ListBox lstCands 
      Height          =   1425
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   4575
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H80000004&
      Height          =   735
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmArezzoDecision.frx":0316
      Top             =   120
      Width           =   4575
   End
   Begin VB.CommandButton cmdCommit 
      Caption         =   "Select option"
      Default         =   -1  'True
      Height          =   495
      Left            =   3420
      TabIndex        =   0
      Top             =   4440
      Width           =   1400
   End
   Begin VB.Label lblCandName 
      Caption         =   "CandName"
      Height          =   270
      Left            =   6390
      TabIndex        =   9
      Top             =   120
      Width           =   3090
   End
   Begin VB.Label Label1 
      Caption         =   "Candidate:"
      Height          =   255
      Left            =   5280
      TabIndex        =   8
      Top             =   120
      Width           =   900
   End
   Begin VB.Label lblPlease 
      Caption         =   "Label1"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   4455
   End
End
Attribute VB_Name = "frmArezzoDecision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------
' File: frmArezzoDecision.frm
' Copyright InferMed Ltd 1999 All Rights Reserved
' Author: Nicky Johns, InferMed
' Purpose: Deal with Arezzo decisions in MACRO Data Management
'-----------------------------------------
' REVISIONS
'   NCJ 28-29 Sep 99 - Initial Development
'   NCJ 1 Nov 99 - Use candidate captions
'   NCJ 5-8 Nov 99 - Implement Explain button
' MACRO 2.2
'   NCJ 1 Oct 01 - Changed RefreshMe to Display
'   NCJ 17 Jan 02 - Added Hourglass suspend/resume (Buglist 2.2.3, Bug 25)
'   NCJ 10 Feb 03 - Use new DEBS routine to commit decision
'-----------------------------------------

Option Explicit

' The decision instance to which this window belongs
Dim moDecision As TaskInstance
' The collection of candidate names
' which match the candidate captions in the list box (NCJ 1 Nov 99)
Dim mcolCandNames As Collection
' The currently selected candidate
Dim msCurrentCand As String
' Store whether "explanation" window is showing
Dim mbExplainShowing As Boolean

Dim mbOKClicked As Boolean

' NCJ 31 Jan 03
Dim moArezzo As Arezzo_DM

'-----------------------------------------
Public Function Display(oTask As TaskInstance, oArezzo As Arezzo_DM) As Boolean
'-----------------------------------------
' Refresh and display for the given decision task.
' Returns TRUE if the user did anything, or FALSE if nothing done
' NCJ 31 Jan 03 - Added oArezzo argument
'-----------------------------------------
    
    On Error GoTo ErrHandler
    
    Set moArezzo = oArezzo
    
    mbOKClicked = False
    msCurrentCand = ""
    
    Set moDecision = oTask
    If FillCandsBox Then
        ' There were some candidates
        Me.Caption = "Decision - " & moDecision.Name
        lblPlease.Caption = "Please select one of the following options:"
        txtDesc.Text = moDecision.Description
'        cmdExplain.Enabled = False
'        cmdCommit.Enabled = False
       ' NCJ 17 Jan 02 - Suspend/resume hourglass
        HourglassSuspend
        Me.Show vbModal
        HourglassResume
    
    End If
    
    Display = mbOKClicked

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

'-----------------------------------------
Private Sub cmdCancel_Click()
'-----------------------------------------

    Unload Me

End Sub

'-----------------------------------------
Private Sub cmdCommit_Click()
'-----------------------------------------
' Commit to the selected option
' NCJ 10 Feb 03 - Use new CommitDecision DEBS call
'-----------------------------------------
Dim sCand As String
    
    On Error GoTo ErrHandler
'    Call moDecision.Commit(msCurrentCand)
    Call moArezzo.CommitDecision(moDecision.TaskKey, msCurrentCand)
    mbOKClicked = True
    
    Unload Me
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdCommit_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'------------------------------------------
Private Sub cmdExplain_Click()
'------------------------------------------
' Provide explanation for a candidate
'------------------------------------------

    ' Display the explanation box
    ShowExplainBox
    
    Call RefreshExplanation(msCurrentCand)

End Sub

'------------------------------------------
Private Sub RefreshExplanation(sCand As String)
'------------------------------------------
' Refresh the explanation window for this candidate
'------------------------------------------
Dim sExplain As String
Dim colExplains As Collection
Dim vExplain As Variant
Dim sReasonsFor As String
Dim sReasonsAgainst As String
Dim sRecommend As String
    
    On Error GoTo ErrHandler

    sExplain = ""
    
    ' Get the arguments for
    Set colExplains = moDecision.Explain(sCand, "for")
    sReasonsFor = ""
    If colExplains.Count > 0 Then
        sReasonsFor = "Reasons in favour: " & vbCrLf
        For Each vExplain In colExplains
            sReasonsFor = sReasonsFor & " " & CStr(vExplain) & vbCrLf
        Next
    End If
    
    ' Get the arguments against
    Set colExplains = moDecision.Explain(sCand, "against")
    sReasonsAgainst = ""
    If colExplains.Count > 0 Then
        sReasonsAgainst = "Reasons against: " & vbCrLf
        For Each vExplain In colExplains
            sReasonsAgainst = sReasonsAgainst & " " & CStr(vExplain) & vbCrLf
        Next
    End If
    
    ' Get the recommendation
    If moDecision.IsRecommended(sCand) Then
        sRecommend = moDecision.RecommendationExplain(sCand)
        If sRecommend > "" Then
            sRecommend = "This candidate is recommended because: " & vbCrLf _
                    & " " & sRecommend
        End If
    End If
    
    ' Now combine the strings
    If sReasonsFor > "" Then
        sExplain = sReasonsFor & vbCrLf
    End If
    If sReasonsAgainst > "" Then
        sExplain = sExplain & sReasonsAgainst & vbCrLf
    End If
    If sRecommend > "" Then
        sExplain = sExplain & sRecommend
    End If
    ' See if there was anything
    If sExplain = "" Then
        sExplain = "No relevant information is available for this candidate"
    End If
    
    txtExplain.Text = sExplain
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "RefreshExplanation")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'------------------------------------------
Private Sub cmdHideExplain_Click()
'------------------------------------------
' Hide the explanation window
'------------------------------------------

    Call HideExplainBox
    
End Sub

'------------------------------------------
Private Sub Form_Load()
'------------------------------------------
' Place the form centrally and hide explanation area
'------------------------------------------
    On Error GoTo ErrHandler
    
    Me.Top = (Screen.Height - Me.Height) \ 2
    Me.Left = (Screen.Width - Me.Width) \ 2
    ' Initially hide the explanation field - NCJ 5 Nov 99
    HideExplainBox
    cmdExplain.Enabled = False
    cmdCommit.Enabled = False
    
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

'------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'------------------------------------------
' Tidy up
'---------------------------------------------------------

    Set moDecision = Nothing
    Set mcolCandNames = Nothing
    Set moArezzo = Nothing
    
End Sub

'------------------------------------------
Private Sub lstCands_Click()
'------------------------------------------
' Click on a candidate
'------------------------------------------
Dim sCand As String
Dim sCandInfo As String

    On Error GoTo ErrHandler

    sCand = GetSelectedItem(lstCands)
    
    If sCand > "" Then     ' Anything selected?
        If sCand = msCurrentCand Then   ' Same as before?
            ' Do nothing
        Else
            msCurrentCand = sCand
            cmdExplain.Enabled = True
            cmdCommit.Enabled = True
            ' Set frame caption
            fraCand.Caption = lstCands.Text     ' The candidate caption
            ' Set label in explanation section
            lblCandName.Caption = lstCands.Text
            sCandInfo = "Netsupport = " & moDecision.NetSupport(sCand) & vbCrLf
            If moDecision.IsRecommended(sCand) Then
                sCandInfo = sCandInfo & "This is a recommended option"
            Else
                sCandInfo = sCandInfo & "This is not a recommended option"
            End If
            txtCandInfo.Text = sCandInfo
            ' Change explanation if explanations showing
            If mbExplainShowing Then
                Call RefreshExplanation(sCand)
            End If
        End If
    Else
        HideExplainBox
        cmdExplain.Enabled = False
        cmdCommit.Enabled = False
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "lstCands_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'------------------------------------------
Private Function FillCandsBox() As Boolean
'------------------------------------------
' Fill the list box with candidates
' Returns TRUE if there's at least one candidate
' Returns FALSE if no candidates
'------------------------------------------
Dim vCand As Variant

    On Error GoTo ErrHandler
    lstCands.Clear
    Set mcolCandNames = New Collection
    If moDecision.Candidates.Count > 0 Then
        For Each vCand In moDecision.Candidates
            ' Add the caption to the list box
            ' and the candidate name to our own collection
            mcolCandNames.Add CStr(vCand)
            lstCands.AddItem moDecision.CandidateCaption(CStr(vCand))
        Next
        ' Select the first candidate
        lstCands.ListIndex = 0
        FillCandsBox = True
    Else
        FillCandsBox = False    ' No candidates! (Shouldn't happen...)
    End If
    
Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "FillCandsBox")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Function

'---------------------------------------------------------
Private Function GetSelectedItem(oListBox As ListBox) As String
'---------------------------------------------------------
' Get the currently selected candidate from the listbox
' NB Return candidate name corresponding to chosen candidate caption
'---------------------------------------------------------
Dim nIndex As Integer
    
    On Error GoTo ErrHandler
    nIndex = oListBox.ListIndex
    If nIndex > -1 Then
        ' ListIndex starts at 0 but collections start at 1
        GetSelectedItem = mcolCandNames(nIndex + 1)
        ' GetSelectedItem = oListBox.List(nIndex)
    Else
        GetSelectedItem = ""
    End If
    
Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "GetSelectedItem")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Function

'---------------------------------------------------------
Private Sub ShowExplainBox()
'---------------------------------------------------------
' Show the explanation box by expanding the window size
'---------------------------------------------------------

    Me.Width = 9885
    mbExplainShowing = True

End Sub

'---------------------------------------------------------
Private Sub HideExplainBox()
'---------------------------------------------------------
' Hide the explanation box by shrinking the window size
'---------------------------------------------------------
    
    Me.Width = 5130
    mbExplainShowing = False

End Sub
