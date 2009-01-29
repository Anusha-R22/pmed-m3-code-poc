VERSION 5.00
Begin VB.Form frmORAMA 
   BackColor       =   &H00CFCFCF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ORAMA CDS"
   ClientHeight    =   12480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10485
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   12480
   ScaleWidth      =   10485
   Begin VB.VScrollBar vsbScroll 
      Height          =   12255
      Left            =   10080
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   330
   End
   Begin VB.PictureBox picORAMAPage 
      BackColor       =   &H00CFCFCF&
      BorderStyle     =   0  'None
      Height          =   12135
      Left            =   120
      ScaleHeight     =   12135
      ScaleWidth      =   9855
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      Begin VB.Frame fraOpts 
         BackColor       =   &H00CFCFCF&
         Caption         =   "Option Frame"
         Height          =   1695
         Index           =   0
         Left            =   6960
         TabIndex        =   7
         Top             =   6960
         Width           =   1575
      End
      Begin VB.CheckBox chkCand 
         BackColor       =   &H00CFCFCF&
         Caption         =   "Stop oral iron"
         ForeColor       =   &H007D00C7&
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   6
         Top             =   6600
         Width           =   4095
      End
      Begin VB.OptionButton optCand 
         BackColor       =   &H00CFCFCF&
         Caption         =   "Not recommended option"
         ForeColor       =   &H007D00C7&
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   5
         Top             =   8280
         Width           =   3375
      End
      Begin VB.PictureBox picNextButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   9120
         ScaleHeight     =   345
         ScaleWidth      =   345
         TabIndex        =   1
         ToolTipText     =   "Save decisions and move to next form"
         Top             =   8880
         Width           =   375
      End
      Begin VB.Label lblRef 
         BackColor       =   &H00CFCFCF&
         Caption         =   "Reference"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   5160
         TabIndex        =   14
         Top             =   5040
         Width           =   1095
      End
      Begin VB.Image imgCancel 
         Height          =   480
         Left            =   600
         Picture         =   "frmORAMA.frx":0000
         ToolTipText     =   "Cancel this form"
         Top             =   11160
         Width           =   405
      End
      Begin VB.Label lblHeading 
         Alignment       =   2  'Center
         BackColor       =   &H00CFCFCF&
         Caption         =   "ORAMA CDS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007D00C7&
         Height          =   375
         Left            =   3480
         TabIndex        =   13
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lblIntro1 
         AutoSize        =   -1  'True
         BackColor       =   &H00CFCFCF&
         Caption         =   $"frmORAMA.frx":0408
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   16860
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblIntro2 
         AutoSize        =   -1  'True
         BackColor       =   &H00CFCFCF&
         Caption         =   $"frmORAMA.frx":04C4
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   16425
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblIntro3 
         AutoSize        =   -1  'True
         BackColor       =   &H00CFCFCF&
         Caption         =   $"frmORAMA.frx":0584
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   17175
         WordWrap        =   -1  'True
      End
      Begin VB.Line lnTopLine 
         X1              =   120
         X2              =   9840
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackColor       =   &H00CFCFCF&
         Caption         =   "Workup Required?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007D00C7&
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   9
         Top             =   5040
         Width           =   2160
      End
      Begin VB.Label lblRecommend 
         BackColor       =   &H00D6C1F5&
         Caption         =   "Recommendation"
         ForeColor       =   &H007D00C7&
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   8
         Top             =   9240
         Width           =   2175
      End
      Begin VB.Shape shpRecBox 
         BackColor       =   &H00D6C1F5&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H007D00C7&
         BorderWidth     =   3
         Height          =   1455
         Index           =   0
         Left            =   600
         Top             =   9000
         Width           =   6495
      End
      Begin VB.Label lblReasons 
         BackColor       =   &H00CFCFCF&
         Caption         =   "Reasons why not:"
         ForeColor       =   &H005E145C&
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   4
         Top             =   7080
         Width           =   1815
      End
      Begin VB.Label lblReason 
         AutoSize        =   -1  'True
         BackColor       =   &H00CFCFCF&
         Caption         =   "Blah blah blah "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005E145C&
         Height          =   240
         Index           =   0
         Left            =   1200
         TabIndex        =   3
         Top             =   7440
         Width           =   1290
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   240
         X2              =   9840
         Y1              =   5880
         Y2              =   5880
      End
   End
End
Attribute VB_Name = "frmORAMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       frmORAMA.frm
'   Copyright:  InferMed Ltd. 2004. All Rights Reserved
'   Author:     Nicky Johns, October 2004
'   Purpose:    Arezzo decision display for ORAMA in Windows DE
'----------------------------------------------------------------------------------------'
' NB This form is only used if MACRO_DM is compiled with ORAMA = 1
'----------------------------------------------------------------------------------------'
' Revisions:
'
' NCJ 26 Oct 04 - Initial version
' NCJ 1-9 Nov 04 - Further development work
'----------------------------------------------------------------------------------------'

Option Explicit

' The colours we'll use
Private Const mlGREY = &HCFCFCF
Private Const mlMAGENTA = &H7D00C7
Private Const mlLTPINK = &HD6C1F5
Private Const mlPURPLE = &H5E145C

Private Const msglLEFT_X = 240
Private Const mnGAP = 120
Private msBULLET As String
 
Private mbSomethingHappened As Boolean
' Max X,Y coords on form
Private msglMaxY As Single
Private msglMaxX As Single
Private msglFrameY As Single
Private mnIndex As Integer
Private msglOurWidth As Single

' Collection of OramaDec objects
Private mcolDecisions As Collection
' Collection of keys we had last time
Private mcolHadLastTime As Collection

' Did the user commit the decisions?
Private Enum WhatHappened
    decsAll
    decsSome
    decsNone
End Enum

'----------------------------------------------------------------------------------------'
Public Function Display(oArezzo As Arezzo_DM, ByVal sglTop As Single, _
                        ByVal sglLeft As Single, _
                        ByVal sglWidth As Single, ByVal sglHeight As Single) As Boolean
'----------------------------------------------------------------------------------------'
' Display ORAMA tasks
' Returns TRUE if the user did anything (i.e. a save is needed)
' or FALSE if nothing happened
' Keeps checking for new tasks
'----------------------------------------------------------------------------------------'
Dim colTasks As Collection

    On Error GoTo ErrHandler

    mbSomethingHappened = False
    Set mcolHadLastTime = New Collection
    
    Set colTasks = oArezzo.GetArezzoTasks
    
    Call HourglassOff
    
    Do While AnyNewTasks(colTasks)
        
        Me.Top = sglTop
        Me.Left = sglLeft
        Me.Width = sglWidth
        Me.Height = sglHeight
        
        Call ResizeTopLabels
        
        ' Set up things
        Call InitialiseThings
    
        Call DisplayTasks(colTasks)
        
        Me.Show vbModal
        
        Call CollectionRemoveAll(colTasks)
        Set colTasks = Nothing
        Call UnloadThings
        
        Set colTasks = oArezzo.GetArezzoTasks
    Loop
    
    Display = mbSomethingHappened
    
    If Not colTasks Is Nothing Then
        Call CollectionRemoveAll(colTasks)
        Set colTasks = Nothing
    End If
    
    Set mcolHadLastTime = Nothing

Exit Function
    
ErrHandler:
    Call Err.Raise(Err.Number, , Err.Description & "|frmORAMA.Display")
End Function

'----------------------------------------------------------------------------------------'
Private Sub DisplayTasks(colTasks As Collection)
'----------------------------------------------------------------------------------------'
' Display all the ORAMA tasks in this window
' We only process Decisions and Actions
'----------------------------------------------------------------------------------------'
Dim oTaskInst As TaskInstance
Dim vCand As Variant
Dim oOramaDec As OramaDec
Dim oFrame As Frame

    On Error GoTo Errlabel
    
    For Each oTaskInst In colTasks
        Call BuildCaption(oTaskInst.Caption)
        Select Case oTaskInst.TaskType
        Case "decision"
            ' Only deal with permitted decisions
            If oTaskInst.TaskState = "permitted" Then
                Set oFrame = BuildFrame
                Set oOramaDec = New OramaDec
                Set oOramaDec.DecisionTask = oTaskInst
                For Each vCand In oTaskInst.Candidates
                    oOramaDec.AddIndex BuildCandidate(oTaskInst, CStr(vCand), oFrame)
                Next
                mcolDecisions.Add oOramaDec
                Set oOramaDec = Nothing
                ' Size frame correctly
                oFrame.Height = msglFrameY
                msglMaxY = msglMaxY + oFrame.Height + mnGAP
            End If
        Case "action"
        
        Case Else
            ' Ignore others for ORAMA
        End Select
        Call BuildLine
    Next
    
    Call BuildOKButton
    
    ' Set the scroll bar
    picORAMAPage.Height = msglMaxY
    vsbScroll.Min = 0
    If picORAMAPage.Height > Me.ScaleHeight Then
        vsbScroll.Max = picORAMAPage.Height - Me.ScaleHeight
        vsbScroll.Enabled = True
        vsbScroll.SmallChange = vsbScroll.Max / 10
        vsbScroll.LargeChange = vsbScroll.Max / 5
    Else
        vsbScroll.Max = 0
        vsbScroll.Enabled = False
    End If
    
Exit Sub
Errlabel:
    Call Err.Raise(Err.Number, , Err.Description & "|frmORAMA.DisplayTasks")
End Sub

'----------------------------------------------------------------------------------------'
Private Sub BuildCaption(ByVal sCaption As String)
'----------------------------------------------------------------------------------------'
' Build a task caption
'----------------------------------------------------------------------------------------'
Dim sFileRef As String

    On Error GoTo Errlabel
    
    Load lblCaption(mnIndex)
    With lblCaption(mnIndex)
        ' Strip off the HTML stuff at the end, keeping the file ref
        .Caption = RemoveHTMLStuff(sCaption, sFileRef)
        .Left = msglLEFT_X
        .Top = msglMaxY
        .Visible = True
        ' Update Y coord
        msglMaxY = msglMaxY + .Height    ' + mnGAP
    End With
    mnIndex = mnIndex + 1
    
    ' Create "Reference" label if appropriate
    If sFileRef > "" Then
        Load lblRef(mnIndex)
        With lblRef(mnIndex)
            ' Store the file ref in the tag
            .Tag = sFileRef
            .Top = msglMaxY
            .Left = msglLEFT_X + msglOurWidth - .Width
            .Visible = True
            .Enabled = True
            msglMaxY = msglMaxY + .Height    ' + mnGAP
        End With
        mnIndex = mnIndex + 1
    End If
    
Exit Sub
Errlabel:
    Call Err.Raise(Err.Number, , Err.Description & "|frmORAMA.BuildCaption")
End Sub

'----------------------------------------------------------------------------------------'
Private Function BuildFrame() As Frame
'----------------------------------------------------------------------------------------'
' Build a frame to hold option buttons
'----------------------------------------------------------------------------------------'

    On Error GoTo Errlabel
    
    Load fraOpts(mnIndex)
    With fraOpts(mnIndex)
        .Top = msglMaxY
        .Left = msglLEFT_X
        .Width = msglOurWidth
        .Visible = True
        .BorderStyle = vbBSNone
    End With
    ' Start the Y coord at the top
    msglFrameY = mnGAP
    Set BuildFrame = fraOpts(mnIndex)
    
    mnIndex = mnIndex + 1
    
Exit Function
Errlabel:
    Call Err.Raise(Err.Number, , Err.Description & "|frmORAMA.BuildFrame")
End Function

'----------------------------------------------------------------------------------------'
Private Function BuildCandidate(oDecn As TaskInstance, ByVal sCand As String, _
                oFrame As Frame) As Integer
'----------------------------------------------------------------------------------------'
' Build a candidate option inside oFrame
' and return its index
'----------------------------------------------------------------------------------------'
Dim oControl As Control

    On Error GoTo Errlabel
    
    If oDecn.IsMultiple Then
        ' Multiple choice decision (check boxes)
        Load chkCand(mnIndex)
        Set oControl = chkCand(mnIndex)
    Else
        ' Single choice decision (option buttons)
        Load optCand(mnIndex)
        Set oControl = optCand(mnIndex)
    End If
    
    ' Return the index value
    BuildCandidate = mnIndex
    
    mnIndex = mnIndex + 1
    
    With oControl
        Set .Container = oFrame
        .Left = msglLEFT_X
        .Caption = oDecn.CandidateCaption(sCand)
        .Tag = sCand
        .Visible = True
        .Enabled = True
    End With
    
    If Not oDecn.IsRecommended(sCand) Then
        ' Not recommended - plonk it straight down
        oControl.Top = msglFrameY
        oControl.BackColor = mlGREY
        ' Update Y coord (within frame)
        msglFrameY = msglFrameY + oControl.Height + mnGAP
        Call BuildArgs(oDecn, sCand, oFrame)
    Else
        ' Build a recommended candidate
        Call BuildRecCand(oDecn, sCand, oControl, oFrame)
    End If
    
Exit Function
Errlabel:
    Call Err.Raise(Err.Number, , Err.Description & "|frmORAMA.BuildCandidate")
    
End Function

'----------------------------------------------------------------------------------------'
Private Sub BuildRecCand(oDecn As TaskInstance, ByVal sCand As String, _
                oCandControl As Control, _
                oFrame As Frame)
'----------------------------------------------------------------------------------------'
' Build a "recommended" candidate inside a pink box
' oCandControl has already been created (checkbox or option button)
'----------------------------------------------------------------------------------------'
Dim nPinkIndex As Integer

    On Error GoTo Errlabel
    
    ' Load the pink box
    nPinkIndex = mnIndex
    Load shpRecBox(nPinkIndex)
    With shpRecBox(nPinkIndex)
        Set .Container = oFrame
        .Visible = True
        .Left = 0
        .Top = msglFrameY
        .Width = msglOurWidth
        msglFrameY = msglFrameY + mnGAP * 2
    End With
    
    mnIndex = mnIndex + 1
        
    ' Add "Recommendation" label
    Load lblRecommend(mnIndex)
    With lblRecommend(mnIndex)
        Set .Container = oFrame
        .Visible = True
        .Left = msglLEFT_X
        .Top = msglFrameY
        .ZOrder
        msglFrameY = msglFrameY + .Height + mnGAP
    End With
    
    mnIndex = mnIndex + 1
    
    ' Place the candidate control
    With oCandControl
        .Top = msglFrameY
        .BackColor = mlLTPINK
        .Width = shpRecBox(nPinkIndex).Width - .Left - mnGAP
        msglFrameY = msglFrameY + .Height + mnGAP
    End With
    
    Call BuildArgs(oDecn, sCand, oFrame)
    
    ' Set the height of the pink box
    shpRecBox(nPinkIndex).Height = msglFrameY - shpRecBox(nPinkIndex).Top
    msglFrameY = msglFrameY + mnGAP
    
Exit Sub
Errlabel:
    Call Err.Raise(Err.Number, , Err.Description & "|frmORAMA.BuildRecCand")
End Sub

'---------------------------------------------------------------------
Private Sub BuildArgs(oDecn As TaskInstance, ByVal sCand As String, _
                oFrame As Frame)
'---------------------------------------------------------------------
' Display the arguments for this decision candidate
'---------------------------------------------------------------------
Dim vArg As Variant
Dim bRecommended As Boolean
    
    On Error GoTo Errlabel
    
    bRecommended = oDecn.IsRecommended(sCand)
    
    If oDecn.Explain(sCand, "for").Count > 0 Then
        Call BuildReasonsLabel("Reasons why:", bRecommended, oFrame)
        For Each vArg In oDecn.Explain(sCand, "for")
            Call BuildArg(CStr(vArg), bRecommended, oFrame)
        Next
    End If

    If oDecn.Explain(sCand, "against").Count > 0 Then
        Call BuildReasonsLabel("Reasons why not:", bRecommended, oFrame)
        For Each vArg In oDecn.Explain(sCand, "against")
            Call BuildArg(CStr(vArg), bRecommended, oFrame)
        Next
    End If

Exit Sub
Errlabel:
    Call Err.Raise(Err.Number, , Err.Description & "|frmORAMA.BuildArgs")
End Sub

'---------------------------------------------------------------------
Private Sub BuildArg(ByVal sArg As String, ByVal bRecommended As Boolean, oFrame As Frame)
'---------------------------------------------------------------------
' Build an argument
'---------------------------------------------------------------------

    On Error GoTo Errlabel
    
    Load lblReason(mnIndex)
    With lblReason(mnIndex)
        Set .Container = oFrame
        .Caption = msBULLET & " " & sArg
        .Top = msglFrameY
        .Left = msglLEFT_X + mnGAP * 3
        .Width = oFrame.Width - .Left - mnGAP
        If bRecommended Then
            .BackColor = mlLTPINK
        Else
            .BackColor = mlGREY
        End If
        ' Make visible and bring to the front
        .Visible = True
        .ZOrder
        msglFrameY = msglFrameY + .Height + mnGAP
    End With
    mnIndex = mnIndex + 1

Exit Sub
Errlabel:
    Call Err.Raise(Err.Number, , Err.Description & "|frmORAMA.BuildArg")
End Sub

'---------------------------------------------------------------------
Private Sub BuildReasonsLabel(ByVal sReasonsText As String, ByVal bRecommended As Boolean, oFrame As Frame)
'---------------------------------------------------------------------
' Build label that says "Reasons for" or "Reasons against"
'---------------------------------------------------------------------

    Load lblReasons(mnIndex)
    With lblReasons(mnIndex)
        Set .Container = oFrame
        .Caption = sReasonsText
        .Top = msglFrameY
        .Left = msglLEFT_X + mnGAP * 2
        If bRecommended Then
            .BackColor = mlLTPINK
        Else
            .BackColor = mlGREY
        End If
        ' Make visible and bring to the front
        .Visible = True
        .ZOrder
        msglFrameY = msglFrameY + .Height + mnGAP
    End With
    mnIndex = mnIndex + 1
    
End Sub

'---------------------------------------------------------------------
Private Sub BuildOKButton()
'---------------------------------------------------------------------
' Put the blue arrow in a suitable position, and the garbage can
'---------------------------------------------------------------------

    On Error GoTo Errlabel
    
    ' Do the garbage
    imgCancel.Top = msglMaxY
    imgCancel.Left = msglLEFT_X
    
    ' Let the "Next" button resize itself
    Set picNextButton.Picture = frmImages.imglistStatus.ListImages("DM30_NextEFormOn").Picture
    
    With picNextButton
        .BorderStyle = vbBSNone
        .Top = msglMaxY
        .Left = lnTopLine.X2 - picNextButton.Width
        .BackColor = mlGREY
        .Visible = True
        .Enabled = True
        .TooltipText = "Save decisions and go to next form"
'        .TabIndex = mnTabOrder
'        .TabStop = Not bReadOnly
        msglMaxY = msglMaxY + .Height + mnGAP
    End With

Exit Sub
Errlabel:
    Err.Raise Err.Number, , Err.Description & "|frmORAMA.BuildOKButton"
End Sub

'---------------------------------------------------------------------
Private Sub BuildLine()
'---------------------------------------------------------------------
' Build a line across
'---------------------------------------------------------------------

    On Error GoTo Errlabel
    
    Load Line1(mnIndex)
    With Line1(mnIndex)
        .X1 = msglLEFT_X
        .X2 = lnTopLine.X2
        .Y1 = msglMaxY
        .Y2 = msglMaxY
        .Visible = True
    End With
    ' Update Y coord and index (leave double gap after a line)
    msglMaxY = msglMaxY + mnGAP * 2
    
    mnIndex = mnIndex + 1
    
Exit Sub
Errlabel:
    Err.Raise Err.Number, , Err.Description & "|frmORAMA.BuildLine"
End Sub

'---------------------------------------------------------------------
Private Sub InitialiseThings()
'---------------------------------------------------------------------
' Initialise stuff prior to displaying the tasks
'---------------------------------------------------------------------

    On Error GoTo Errlabel
    
    picORAMAPage.Top = 0
    picORAMAPage.Left = 0
    
    msglMaxY = lnTopLine.Y1 + 2 * mnGAP
    msglMaxX = msglLEFT_X
    picORAMAPage.Width = msglLEFT_X + msglOurWidth + mnGAP
    
    ' Position the scroll bar
    vsbScroll.Top = picORAMAPage.Top
    vsbScroll.Left = picORAMAPage.Left + picORAMAPage.Width
    vsbScroll.Height = Me.ScaleHeight
    
    mnIndex = 1
    Set mcolDecisions = New Collection
  
  ' Hide and disable all the "seed" controls
    Line1(0).Visible = False
    shpRecBox(0).Visible = False
    fraOpts(0).Visible = False
    lblRecommend(0).Visible = False
    lblRef(0).Visible = False
    lblRef(0).FontUnderline = True
    lblCaption(0).Visible = False
    lblCaption(0).Width = msglOurWidth
    lblReasons(0).Visible = False
    lblReason(0).Visible = False
    lblReason(0).Width = msglOurWidth - msglLEFT_X
    optCand(0).Visible = False
    optCand(0).Enabled = False
    optCand(0).Width = msglOurWidth - msglLEFT_X
    chkCand(0).Visible = False
    chkCand(0).Enabled = False
    chkCand(0).Width = msglOurWidth - msglLEFT_X
  
Exit Sub
Errlabel:
    Err.Raise Err.Number, , Err.Description & "|frmORAMA.InitialiseThings"
End Sub

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------
'---------------------------------------------------------------------
    
    msBULLET = "*"
    ' There weren't any before
    Set mcolHadLastTime = New Collection

End Sub

'---------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'---------------------------------------------------------------------

    Call UnloadThings
    
End Sub

'---------------------------------------------------------------------
Private Sub imgCancel_Click()
'---------------------------------------------------------------------
' They want to cancel
'---------------------------------------------------------------------
Dim sMsg As String
Dim bToMoveOn As Boolean

    sMsg = ""
    bToMoveOn = True
    
    ' Did they do anything?
    Select Case UserActivity
    Case WhatHappened.decsNone
        ' Nothing to save
    Case WhatHappened.decsSome, WhatHappened.decsAll
        sMsg = "Are you sure you wish to cancel without saving your choices?"
    End Select
    
    If sMsg > "" Then
        bToMoveOn = (DialogQuestion(sMsg, "ORAMA Clinical Decision Support") = vbYes)
    End If
    
    If bToMoveOn Then
        mbSomethingHappened = False
        Me.Hide
    End If

End Sub

'---------------------------------------------------------------------
Private Sub lblRef_Click(Index As Integer)
'---------------------------------------------------------------------
' Display reference text for a decision
' The file name is in the tag of the label
'---------------------------------------------------------------------

'    Call MsgBox("Display reference " & lblRef(Index).Tag)
    If lblRef(Index).Tag > "" Then
        Call ShowReference(App.Path & "\ORAMA Guidelines\" & lblRef(Index).Tag, "ORAMA")
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub picNextButton_Click()
'---------------------------------------------------------------------
' Move along now
'---------------------------------------------------------------------
Dim sMsg As String
Dim bToSave As Boolean
Dim bToMoveOn As Boolean

    sMsg = ""
    bToSave = True
    bToMoveOn = True
    
    ' Did they do anything?
    Select Case UserActivity
    Case WhatHappened.decsNone
        sMsg = "No choices have been made."
        ' Nothing to save
        bToSave = False
    Case WhatHappened.decsSome
        sMsg = "Not all choices have been made."
    Case WhatHappened.decsAll
        ' That's OK
    End Select
    
    If sMsg > "" Then
        sMsg = sMsg & vbCrLf & "Are you sure you wish to move to the next form?"
        bToMoveOn = (DialogQuestion(sMsg, "ORAMA Clinical Decision Support") = vbYes)
    End If
    
    ' Don't save if they're not moving on
    If bToSave And bToMoveOn Then
        mbSomethingHappened = SaveDecisions
    End If
    
    If bToMoveOn Then
        ' Move on
        Me.Hide
    End If
        
End Sub

'---------------------------------------------------------------------
Private Function SaveDecisions() As Boolean
'---------------------------------------------------------------------
' Process what they've done and save the decision results to AREZZO
' Returns TRUE if any decisions were committed
'---------------------------------------------------------------------
Dim bChanged As Boolean
Dim oOramaDec As OramaDec
Dim vIndex As Variant
Dim colCands As Collection
Dim bRecommit As Boolean
Dim sCand As String

    On Error GoTo ErrHandler
    
    bChanged = False
    
    For Each oOramaDec In mcolDecisions
        If oOramaDec.DecisionTask.IsMultiple Then
            ' Multiple choice
            Set colCands = New Collection
            bRecommit = False
            For Each vIndex In oOramaDec.IndexCollection
                If chkCand(CInt(vIndex)).Value = 1 Then
                    ' Add this candidate to the collection
                    Call colCands.Add(chkCand(CInt(vIndex)).Tag)
                    ' Will need to do it again for non-recommended cand
                    If Not oOramaDec.DecisionTask.IsRecommended(chkCand(CInt(vIndex)).Tag) Then
                        bRecommit = True
                    End If
                End If
            Next
            If colCands.Count > 0 Then
                ' Commit the selected candidates
                Call oOramaDec.DecisionTask.CommitCandidates(colCands)
                bChanged = True
                ' Remember we've done it
                oOramaDec.IsCommitted = True
                If bRecommit Then
                    Call oOramaDec.DecisionTask.CommitCandidates(colCands)
                End If
            End If
        Else
            ' Single choice
            For Each vIndex In oOramaDec.IndexCollection
                If optCand(CInt(vIndex)).Value = True Then
                    ' Retrieve the candidate name
                    sCand = optCand(CInt(vIndex)).Tag
                    Call oOramaDec.DecisionTask.Commit(sCand)
                    bChanged = True
                    ' Remember we've done it
                    oOramaDec.IsCommitted = True
                    If Not oOramaDec.DecisionTask.IsRecommended(sCand) Then
                    ' Need to do it again for non-recommended cand
                        Call oOramaDec.DecisionTask.Commit(sCand)
                    End If
                    ' There can only be ONE selected so we stop now
                    Exit For
                End If
            Next
        End If
    Next

    SaveDecisions = bChanged

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmORAMA.SaveDecisions"

End Function

'---------------------------------------------------------------------
Private Function UserActivity() As WhatHappened
'---------------------------------------------------------------------
' Process what they've done
' Returns TRUE if any decisions were committed
'---------------------------------------------------------------------
Dim oOramaDec As OramaDec
Dim vIndex As Variant
Dim nDoneCount As Integer

    On Error GoTo ErrHandler
    
    nDoneCount = 0
    
    For Each oOramaDec In mcolDecisions
        For Each vIndex In oOramaDec.IndexCollection
            If oOramaDec.DecisionTask.IsMultiple Then
                ' Multiple choice
                If chkCand(CInt(vIndex)).Value = 1 Then
                    ' They chose something
                    nDoneCount = nDoneCount + 1
                    Exit For    ' this decision
                End If
            Else
                ' Single choice
                If optCand(CInt(vIndex)).Value = True Then
                    ' They chose something
                    nDoneCount = nDoneCount + 1
                    Exit For    ' this decision
                End If
            End If
        Next
    Next

    ' Decide what to return
    If nDoneCount = mcolDecisions.Count Then
        UserActivity = decsAll
    ElseIf nDoneCount = 0 Then
        UserActivity = decsNone
    Else
        UserActivity = decsSome
    End If

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmORAMA.UserActivity"

End Function

'---------------------------------------------------------------------
Private Sub UnloadThings()
'---------------------------------------------------------------------
' Unload stuff prior to closing the window
'---------------------------------------------------------------------
Dim oControl As Control
Dim oDec As OramaDec

    On Error GoTo Errlabel
    
    ' Ensure that all controls used for display are unloaded
    ' (except the base element of each control array whose Index is 0)
    For Each oControl In Me.Controls
        If oControl.Name = "Line1" _
        Or oControl.Name = "shpRecBox" _
        Or oControl.Name = "lblRecommend" _
        Or oControl.Name = "lblCaption" _
        Or oControl.Name = "lblReason" _
        Or oControl.Name = "lblReasons" _
        Or oControl.Name = "lblRef" _
        Or oControl.Name = "optCand" _
        Or oControl.Name = "chkCand" Then
            If oControl.Index > 0 Then
                Unload oControl
            End If
        End If
    Next
    
    For Each oControl In Me.Controls
        If oControl.Name = "fraOpts" Then
            If oControl.Index > 0 Then
                Unload oControl
            End If
        End If
    Next
    
    ' Tidy up collections and remember what we had left
    Set mcolHadLastTime = New Collection
    If Not mcolDecisions Is Nothing Then
        For Each oDec In mcolDecisions
            If Not oDec.IsCommitted Then
                ' Remember this is still hanging around
                mcolHadLastTime.Add oDec.DecisionTask.TaskKey, str(oDec.DecisionTask.TaskKey)
            End If
            Call oDec.Terminate
        Next
        Call CollectionRemoveAll(mcolDecisions)
        Set mcolDecisions = Nothing
    End If

Exit Sub
Errlabel:
    Err.Raise Err.Number, , Err.Description & "|frmORAMA.UnloadThings"

End Sub

'---------------------------------------------------------------------
Private Sub vsbScroll_Change()
'---------------------------------------------------------------------
' Vertical scroll bar click
'---------------------------------------------------------------------
    
    picORAMAPage.Top = CSng(-vsbScroll.Value)

End Sub

'---------------------------------------------------------------------
Private Sub vsbScroll_Scroll()
'---------------------------------------------------------------------
' Scroll continuously
'---------------------------------------------------------------------
    
    picORAMAPage.Top = CSng(-vsbScroll.Value)

End Sub

'---------------------------------------------------------------------
Private Function RemoveHTMLStuff(ByVal sCaption As String, ByRef sRefFile As String) As String
'---------------------------------------------------------------------
' Remove the HTML info from a task caption
' Assume we want everything up to the first "<"
'---------------------------------------------------------------------
Dim nPos As Integer

' The ref we want follows on from this
Const sREFTAG As String = "window.open('"

    nPos = InStr(1, sCaption, "<")
    If nPos > 0 Then
        RemoveHTMLStuff = Mid(sCaption, 1, nPos - 1)
    Else
        RemoveHTMLStuff = sCaption
    End If
    
    nPos = InStr(1, sCaption, sREFTAG)
    If nPos > 0 Then
        sRefFile = Mid(sCaption, nPos + Len(sREFTAG))
        nPos = InStr(1, sRefFile, "'")
        If nPos > 0 Then
            sRefFile = Mid(sRefFile, 1, nPos - 1)
        Else
            sRefFile = ""
        End If
    Else
        sRefFile = ""
    End If
    
End Function

'---------------------------------------------------------------------
Private Sub ResizeTopLabels()
'---------------------------------------------------------------------
' Resize the top set of fixed text labels and the line across
'---------------------------------------------------------------------

    ' Centre the heading
    lblHeading.Left = (Me.ScaleWidth - lblHeading.Width) / 2
    
    ' The width we'll use
    msglOurWidth = Me.ScaleWidth - msglLEFT_X - mnGAP - vsbScroll.Width
    
    lblIntro1.Width = msglOurWidth
    lblIntro1.Left = msglLEFT_X
    lblIntro2.Width = msglOurWidth
    lblIntro2.Left = msglLEFT_X
    lblIntro2.Top = lblIntro1.Top + lblIntro1.Height + mnGAP
    lblIntro3.Width = msglOurWidth
    lblIntro3.Left = msglLEFT_X
    lblIntro3.Top = lblIntro2.Top + lblIntro2.Height + mnGAP
    
    lnTopLine.Y1 = lblIntro3.Top + lblIntro3.Height + mnGAP
    lnTopLine.Y2 = lnTopLine.Y1
    lnTopLine.X1 = msglLEFT_X
    lnTopLine.X2 = lnTopLine.X1 + msglOurWidth
    
End Sub

'---------------------------------------------------------------------
Private Sub ShowReference(ByVal sFileName As String, ByVal sTitle As String)
'---------------------------------------------------------------------
' Show decision reference stuff
' sFileName contains the HTML to be displayed
' All errors are ignored in this routine
'---------------------------------------------------------------------
Dim ofrmRef As frmWebNonMDI
Dim sHTML As String

Const sCLOSE_BUTTON = ">Close<"

    On Error GoTo Errlabel
    
    If Not FileExists(sFileName) Then
        Call DialogInformation("Sorry - this reference is not available", "ORAMA")
        Exit Sub
    End If
    
    Set ofrmRef = New frmWebNonMDI
    
    ' Read in the HTML
    sHTML = StringFromFile(sFileName)
    ' We don't want a Close button here
    sHTML = Replace(sHTML, sCLOSE_BUTTON, "><")
    With ofrmRef
        ' Offset it slightly
        .Top = Me.Top + mnGAP
        .Left = Me.Left + mnGAP
        .Width = msglOurWidth
        .Height = Me.Height / 2     ' Make it half our height
        .Display wdtHTML, sHTML, "auto", True, sTitle
    End With

    Set ofrmRef = Nothing

Exit Sub
Errlabel:
    ' Ignore errors here!
    Debug.Print "ERROR! ", Err.Number, Err.Description & "|frmORAMA.ShowReference"
'    Err.Raise Err.Number, , Err.Description & "|frmORAMA.ShowReference"

End Sub

'---------------------------------------------------------------------
Private Function AnyNewTasks(colTasks As Collection) As Boolean
'---------------------------------------------------------------------
' Any new tasks in this collection that we didn't have last time?
' Compare with HadLastTime collection
'---------------------------------------------------------------------
Dim bNewTasks As Boolean
Dim oTask As TaskInstance

    AnyNewTasks = False
    
    ' No new tasks
    If colTasks.Count = 0 Then Exit Function
    
    AnyNewTasks = True
    
    ' No tasks last time
    If mcolHadLastTime Is Nothing Then Exit Function
    
    If mcolHadLastTime.Count = 0 Then Exit Function
    
    ' More this time than last time
    If colTasks.Count > mcolHadLastTime.Count Then Exit Function
    
    ' OK, so there were tasks last time and there are tasks this time
    For Each oTask In colTasks
        ' Not one we had last time?
        If Not CollectionMember(mcolHadLastTime, str(oTask.TaskKey), False) Then Exit Function
    Next
    
    ' We had all of colTasks last time!
    AnyNewTasks = False
    
End Function
