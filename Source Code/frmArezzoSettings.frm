VERSION 5.00
Begin VB.Form frmArezzoSettings 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3435
   ClientLeft      =   1275
   ClientTop       =   810
   ClientWidth     =   7800
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   6540
      TabIndex        =   10
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   6540
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6540
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6540
      TabIndex        =   7
      Top             =   150
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   30
      TabIndex        =   11
      Top             =   30
      Width           =   6435
      Begin VB.TextBox txtOutputBuffer 
         Height          =   285
         Left            =   1700
         TabIndex        =   6
         Top             =   2740
         Width           =   975
      End
      Begin VB.TextBox txtInputBuffer 
         Height          =   285
         Left            =   1700
         TabIndex        =   5
         Top             =   2380
         Width           =   975
      End
      Begin VB.TextBox txtHeapSpace 
         Height          =   285
         Left            =   1700
         TabIndex        =   4
         Top             =   2020
         Width           =   975
      End
      Begin VB.TextBox txtBacktrackSpace 
         Height          =   285
         Left            =   1700
         TabIndex        =   3
         Top             =   1660
         Width           =   975
      End
      Begin VB.TextBox txtLocalSpace 
         Height          =   285
         Left            =   1700
         TabIndex        =   2
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtTextSpace 
         Height          =   285
         Left            =   1700
         TabIndex        =   1
         Top             =   940
         Width           =   975
      End
      Begin VB.TextBox txtProgramSpace 
         Height          =   285
         Left            =   1700
         TabIndex        =   0
         Top             =   580
         Width           =   975
      End
      Begin VB.Label lblAbMaxOutputBuffer 
         Height          =   255
         Left            =   5160
         TabIndex        =   36
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label lblAbMaxInputBuffer 
         Height          =   255
         Left            =   5160
         TabIndex        =   35
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblAbMaxHeapSpace 
         Height          =   255
         Left            =   5160
         TabIndex        =   34
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblAbMaxBackTrackSpace 
         Height          =   255
         Left            =   5160
         TabIndex        =   33
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblAbMaxLocalSpace 
         Height          =   255
         Left            =   5160
         TabIndex        =   32
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblAbMaxTextSpace 
         Height          =   255
         Left            =   5160
         TabIndex        =   31
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblAbMaxProgramSpace 
         Height          =   255
         Left            =   5160
         TabIndex        =   30
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblabsoluteMax 
         Caption         =   "Maximum Value"
         Height          =   255
         Left            =   5040
         TabIndex        =   29
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "AREZZO Setting"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   195
         Width           =   1335
      End
      Begin VB.Label lblDefOutputBuffer 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   3000
         TabIndex        =   27
         Top             =   2760
         Width           =   1875
      End
      Begin VB.Label lblDefInputBuffer 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   3000
         TabIndex        =   26
         Top             =   2400
         Width           =   1875
      End
      Begin VB.Label lblDefHeapSpace 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   3000
         TabIndex        =   25
         Top             =   2040
         Width           =   1875
      End
      Begin VB.Label lblDefBacktrackSpace 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   3000
         TabIndex        =   24
         Top             =   1680
         Width           =   1875
      End
      Begin VB.Label lblDefLocalSpace 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   3000
         TabIndex        =   23
         Top             =   1320
         Width           =   1875
      End
      Begin VB.Label lblDefTextSpace 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   3000
         TabIndex        =   22
         Top             =   960
         Width           =   1875
      End
      Begin VB.Label lblDefProgramSpace 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   3000
         TabIndex        =   21
         Top             =   600
         Width           =   1875
      End
      Begin VB.Label lblRange 
         Caption         =   "Recommended Range"
         Height          =   255
         Left            =   3240
         TabIndex        =   20
         Top             =   180
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Stored Value (K)"
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   195
         Width           =   1335
      End
      Begin VB.Label lblOutBuffer 
         Caption         =   "Output Buffer"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   2740
         Width           =   1455
      End
      Begin VB.Label lblInpBuffer 
         Caption         =   "Input Buffer"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2380
         Width           =   1455
      End
      Begin VB.Label lblHSpace 
         Caption         =   "Heap Space"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2020
         Width           =   1455
      End
      Begin VB.Label lblBSpace 
         Caption         =   "Backtrack Space"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1660
         Width           =   1455
      End
      Begin VB.Label lblLSpace 
         Caption         =   "Local Space"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1300
         Width           =   1455
      End
      Begin VB.Label lblTSpace 
         Caption         =   "Text Space"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   940
         Width           =   1455
      End
      Begin VB.Label lblPSpace 
         Caption         =   "Program Space"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   580
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmArezzoSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2002-2006. All Rights Reserved
'   File:       frmArezzoSettings.frm
'   Author:     Zulfiqar Ahmed, 2002
'   Purpose:    To allow the user to view/edit Prolog memory settings
'----------------------------------------------------------------------------------------'
'Revisions:
'ASH 12/9/2002 - Registry keys replaced with calls to new Settings file
'ASH 24/12/2002 - Added recommended ranges
' NCJ 28 Jan 03 - Pass DB connection string into clsAREZZOMemory
' NCJ 13 Jun 06 - Use Study Access Mode to determine editability
'-----------------------------------------------------------------------------------------

Option Explicit

'values allowed in text boxes
Private Const msVALID_NUMBERS = "0123456789" & vbBack & vbCr
'local variable for keeping save state of values
Private mbValuesSaved As Boolean
Private mconMACRO As ADODB.Connection
'Ash 22/11/2002
Private moArezzoMemory As clsAREZZOMemory
Private mbOKClicked As Boolean
Private msMessage As String
Private mbAnythingSaved As Boolean
Private mbValueChanged As Boolean

'---------------------------------------------------------------------
Private Sub SaveValues()
'---------------------------------------------------------------------
'save values in the registry
'ASH 12/9/2002 - Registry keys replaced with calls to new Settings file
'---------------------------------------------------------------------

    On Error GoTo Errorlabel
    
    With moArezzoMemory
        .ProgramSpace = Val(Trim(txtProgramSpace.Text))
        .HeapSpace = Val(Trim(txtHeapSpace.Text))
        .InputSpace = Val(Trim(txtInputBuffer.Text))
        .OutputSpace = Val(Trim(txtOutputBuffer.Text))
        .TextSpace = Val(Trim(txtTextSpace.Text))
        .LocalSpace = Val(Trim(txtLocalSpace.Text))
        .BacktrackSpace = Val(Trim(txtBacktrackSpace.Text))
    End With
    ' NCJ 28 Jan 03 - Pass in DB connection string
    moArezzoMemory.SaveValues (goUser.CurrentDBConString)
    mbValuesSaved = True
    mbValueChanged = False
    mbAnythingSaved = True
    
    
Exit Sub
Errorlabel:
    mbValuesSaved = False
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "SaveValues", Err.Source) = Retry Then
        Resume
    End If
        
End Sub

'---------------------------------------------------------------------
Private Sub DisplayPrologValues(ByVal enAccessMode As eSDAccessMode)
'---------------------------------------------------------------------
' Load Prolog values from the registry into the form
' NCJ 13 Jun 06 - If not RW, disable text fields
'---------------------------------------------------------------------

    On Error GoTo Errorlabel
    
    'load values into the text boxes
    With moArezzoMemory
        txtProgramSpace.Text = .ProgramSpace
        txtTextSpace.Text = .TextSpace
        txtLocalSpace.Text = .LocalSpace
        txtBacktrackSpace.Text = .BacktrackSpace
        txtHeapSpace.Text = .HeapSpace
        txtInputBuffer.Text = .InputSpace
        txtOutputBuffer.Text = .OutputSpace
    End With
    'load values into tags
    txtProgramSpace.Tag = txtProgramSpace.Text
    txtTextSpace.Tag = txtTextSpace.Text
    txtLocalSpace.Tag = txtLocalSpace.Text
    txtBacktrackSpace.Tag = txtBacktrackSpace.Text
    txtHeapSpace.Tag = txtHeapSpace.Text
    txtInputBuffer.Tag = txtInputBuffer.Text
    txtOutputBuffer.Tag = txtOutputBuffer.Text
    
    ' NCJ 13 Jun 06 - Can only edit if R/W or full control
    If enAccessMode < sdReadWrite Then
        txtProgramSpace.Enabled = False
        txtTextSpace.Enabled = False
        txtLocalSpace.Enabled = False
        txtBacktrackSpace.Enabled = False
        txtHeapSpace.Enabled = False
        txtInputBuffer.Enabled = False
        txtOutputBuffer.Enabled = False
    End If
    
Exit Sub
Errorlabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "DisplayPrologValues", Err.Source) = Retry Then
        Resume
    End If

End Sub

'---------------------------------------------------------------------
Private Sub cmdApply_Click()
'---------------------------------------------------------------------
'save values only if there is a change.
'---------------------------------------------------------------------

    'save values only if a value has been changed
    If mbValueChanged Then
        SaveValues
        cmdApply.Enabled = False
    End If
End Sub

'---------------------------------------------------------------------
Private Sub cmdCancel_Click()
'---------------------------------------------------------------------
'cancel event
'---------------------------------------------------------------------
    
    Unload Me
        
End Sub

'---------------------------------------------------------------------
Private Sub cmdOK_Click()
'---------------------------------------------------------------------
'save values only if there is a change.
'---------------------------------------------------------------------
    
    mbOKClicked = True
    'save values only if a value has been changed
    If mbValueChanged Then
        SaveValues
    End If
    
    Unload Me

End Sub
'---------------------------------------------------------------------
Private Sub cmdReset_Click()
'---------------------------------------------------------------------
'loads default values
'---------------------------------------------------------------------
Dim sMSG As String

    sMSG = "Are you sure you want to reset the settings to their default values?"
    
    If DialogQuestion(sMSG) = vbYes Then
        txtProgramSpace.Text = glPROGRAM_SPACE
        txtTextSpace.Text = glTEXT_SPACE
        txtLocalSpace.Text = glLOCAL_SPACE
        txtBacktrackSpace.Text = glBACKTRACK_SPACE
        txtHeapSpace.Text = glHEAP_SPACE
        txtInputBuffer.Text = glINPUT_SPACE
        txtOutputBuffer.Text = glOUTPUT_SPACE
    End If

End Sub


'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------
    
    Me.Icon = frmMenu.Icon
    FormCentre Me
    
End Sub

'---------------------------------------------------------------------
Private Sub SetTextValue(nKey As Integer)
'---------------------------------------------------------------------
'This procedure will allow only numeric values
'---------------------------------------------------------------------
Dim sChar As String

    On Error GoTo Errorlabel
    
    sChar = Chr(nKey)
    
    If InStr(msVALID_NUMBERS, sChar) = 0 Then
        nKey = 0
    End If
    
    Exit Sub
    
Errorlabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "SetTextValue", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'---------------------------------------------------------------------
'call this event when user is trying to close this window
'---------------------------------------------------------------------
Dim nDialogResult As Integer

    If UnloadMode = vbFormControlMenu Then
        'only ask to save if we have unsaved changed value
        If mbValueChanged And Not mbValuesSaved Then
    
            nDialogResult = DialogQuestion("Some values have been changed. Do you want to save them?", , True)
        
        Select Case nDialogResult
            Case vbYes
                Call SaveValues
            Case vbCancel
                Cancel = 1
        End Select
        
        End If
    End If
    
    If msMessage <> "" And mbAnythingSaved Then
        Call DialogInformation(msMessage, "AREZZO Memory Updates")
    End If

    
End Sub

'---------------------------------------------------------------------
Private Sub txtBacktrackSpace_Change()
'---------------------------------------------------------------------
'checks maximum value
'---------------------------------------------------------------------
    
    ChangeEvent txtBacktrackSpace, glBACKTRACK_SPACE, glMAX_BACKTRACK_SPACE
End Sub

'---------------------------------------------------------------------
Private Sub txtBacktrackSpace_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
'key press event
'---------------------------------------------------------------------
    SetTextValue KeyAscii
End Sub
'---------------------------------------------------------------------
Private Sub txtBacktrackSpace_LostFocus()
'---------------------------------------------------------------------
'checks if the value is smaller than the minimum value allowed
'also checks if the value is greater than the recommended average figure
'---------------------------------------------------------------------
    LostFocus txtBacktrackSpace, glBACKTRACK_SPACE, glMAX_BACKTRACK_SPACE, _
    glMID_BACKTRACK_SPACE, gsBACKTRACK_SPACE, lblBSpace
End Sub

'---------------------------------------------------------------------
Private Sub txtHeapSpace_Change()
'---------------------------------------------------------------------
'checks maximum value
'---------------------------------------------------------------------

    ChangeEvent txtHeapSpace, glHEAP_SPACE, glMAX_HEAP_SPACE
End Sub

'---------------------------------------------------------------------
Private Sub txtHeapSpace_keypress(KeyAscii As Integer)
'---------------------------------------------------------------------
'key press event
'---------------------------------------------------------------------
    SetTextValue KeyAscii
End Sub

'---------------------------------------------------------------------
Private Sub txtHeapSpace_LostFocus()
'---------------------------------------------------------------------
'checks if a value is smaller than the minimum value allowed
'also checks if the value is greater than the recommended average figure
'---------------------------------------------------------------------
    LostFocus txtHeapSpace, glHEAP_SPACE, glMAX_HEAP_SPACE, _
    glMID_HEAP_SPACE, gsHEAP_SPACE, lblHSpace
End Sub

'---------------------------------------------------------------------
Private Sub txtInputBuffer_Change()
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------
    
    ChangeEvent txtInputBuffer, glINPUT_SPACE, glMAX_INPUT_SPACE
End Sub

'---------------------------------------------------------------------
Private Sub txtInputBuffer_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
'key press event
'---------------------------------------------------------------------
    SetTextValue KeyAscii
End Sub
'---------------------------------------------------------------------
Private Sub txtInputBuffer_LostFocus()
'---------------------------------------------------------------------
'checks if the value is smaller than the minimum value allowed
'also checks if the value is greater than the recommended average figure
'---------------------------------------------------------------------
    
    LostFocus txtInputBuffer, glINPUT_SPACE, glMAX_INPUT_SPACE, _
    glMID_INPUT_SPACE, gsINPUT_SPACE, lblInpBuffer
End Sub

'---------------------------------------------------------------------
Private Sub txtLocalSpace_Change()
'---------------------------------------------------------------------
'checks maximum value
'---------------------------------------------------------------------
    
    ChangeEvent txtLocalSpace, glLOCAL_SPACE, glMAX_LOCAL_SPACE
End Sub

'---------------------------------------------------------------------
Private Sub txtLocalSpace_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
'key press event
'---------------------------------------------------------------------
    SetTextValue KeyAscii
End Sub
'---------------------------------------------------------------------
Private Sub txtLocalSpace_LostFocus()
'---------------------------------------------------------------------
'checks if the value is smaller than the minimum value allowed
'also checks if the value is greater than the recommended average figure
'---------------------------------------------------------------------
    
    LostFocus txtLocalSpace, glLOCAL_SPACE, glMAX_LOCAL_SPACE, _
    glMID_LOCAL_SPACE, gsLOCAL_SPACE, lblLSpace
End Sub

'---------------------------------------------------------------------
Private Sub txtOutputBuffer_Change()
'---------------------------------------------------------------------
'checks maximum value
'---------------------------------------------------------------------

    ChangeEvent txtOutputBuffer, glOUTPUT_SPACE, glMAX_OUTPUT_SPACE
End Sub

'---------------------------------------------------------------------
Private Sub txtOutputBuffer_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
'key press event
'---------------------------------------------------------------------
    SetTextValue KeyAscii
End Sub
'---------------------------------------------------------------------
Private Sub txtOutputBuffer_LostFocus()
'---------------------------------------------------------------------
'checks if the value is smaller than the minimum value allowed
'also checks if the value is greater than the recommended average figure
'---------------------------------------------------------------------
    
    LostFocus txtOutputBuffer, glOUTPUT_SPACE, glMAX_OUTPUT_SPACE, _
    glMID_OUTPUT_SPACE, gsOUTPUT_SPACE, lblOutBuffer
End Sub

'---------------------------------------------------------------------
Private Sub txtProgramSpace_Change()
'---------------------------------------------------------------------
'checks maximum value
'---------------------------------------------------------------------

    ChangeEvent txtProgramSpace, glPROGRAM_SPACE, glMAX_PROGRAM_SPACE
    
End Sub

'---------------------------------------------------------------------
Private Sub txtProgramSpace_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
'Key press event
'---------------------------------------------------------------------
    SetTextValue KeyAscii
        
End Sub
'---------------------------------------------------------------------
Private Sub txtProgramSpace_LostFocus()
'---------------------------------------------------------------------
'checks if the vlaue is smaller than the minimum vlaue allowed
'also checks if the value is greater than the recommended average figure
'---------------------------------------------------------------------

    LostFocus txtProgramSpace, glPROGRAM_SPACE, glMAX_PROGRAM_SPACE, _
    glMID_PROGRAM_SPACE, gsPROGRAM_SPACE, lblPSpace
    
End Sub

'---------------------------------------------------------------------
Private Sub txtTextSpace_Change()
'---------------------------------------------------------------------
'checks maximum value
'---------------------------------------------------------------------

    ChangeEvent txtTextSpace, glTEXT_SPACE, glMAX_TEXT_SPACE
End Sub

'---------------------------------------------------------------------
Private Sub txtTextSpace_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
'key press even
'---------------------------------------------------------------------
    SetTextValue KeyAscii

End Sub

'---------------------------------------------------------------------
Private Sub txtTextSpace_LostFocus()
'---------------------------------------------------------------------
'checks if the value is smaller than the minimum value allowed
'also checks if the value is greater than the recommended average figure
'---------------------------------------------------------------------
    
    LostFocus txtTextSpace, glTEXT_SPACE, glMAX_TEXT_SPACE, _
    glMID_TEXT_SPACE, gsTEXT_SPACE, lblTSpace
End Sub

'---------------------------------------------------------------------
Private Sub LostFocus(oTextBox As TextBox, _
                        lDefaultValue As Long, _
                        lMaxValue As Long, _
                        lMIDValue As Long, _
                        sAREZZOMemory As String, _
                        sLabel As String)
'---------------------------------------------------------------------
'generic lost focus event for all text boxes
'---------------------------------------------------------------------
 Dim sMSG As String
    
    On Error GoTo Errorlabel
    
    'do no checks if cancel clicked
    If Me.ActiveControl.Name = "cmdCancel" Then
        Exit Sub
    End If
    
    If Val(oTextBox.Text) < lDefaultValue Then
        DialogInformation "Please choose a value between " & lDefaultValue & " and " & lMaxValue
            Call ResetTextBox(oTextBox, sAREZZOMemory)
        oTextBox.SetFocus
    ElseIf Val(oTextBox.Text) > lMIDValue Then
        sMSG = "The value entered for " & sLabel & " is outside the range " & vbCrLf _
        & " recommended for normal AREZZO use, and may cause" & vbCrLf _
        & " problems on computers without enough memory." & vbCrLf _
        & " Are you sure you want to set this value?"
        If DialogQuestion(sMSG) = vbNo Then
            Call ResetTextBox(oTextBox, sAREZZOMemory)
        End If
    End If

    Exit Sub
    
Errorlabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "LostFocus", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Private Sub ChangeEvent(oTextBox As TextBox, _
                        lDefaultValue As Long, _
                        lMaxValue As Long)
'---------------------------------------------------------------------
'generic Change event for all text boxes
'---------------------------------------------------------------------
    
    On Error GoTo Errorlabel
    
    If Val(oTextBox.Text) > lMaxValue Then
        DialogInformation "Please choose a value between " & lDefaultValue & " and " & lMaxValue
        oTextBox.Text = oTextBox.Tag
    ElseIf Not IsNumeric(oTextBox.Text) Then
        oTextBox.Text = oTextBox.Tag
    Else
        oTextBox.Tag = oTextBox.Text
    End If
    
    mbValueChanged = True
    mbValuesSaved = False
    cmdApply.Enabled = mbValueChanged
    
    Exit Sub
    
Errorlabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "ChangeEvent", Err.Source) = Retry Then
        Resume
    End If
    
End Sub

'--------------------------------------------------------------------------
Private Sub DisplayMinMaxValues()
'--------------------------------------------------------------------------
'ASH 20/11/2002 loads the minimum and maximum values
'--------------------------------------------------------------------------
    'minimum /recommended range values
    lblDefBacktrackSpace = glBACKTRACK_SPACE & Space(1) & "-" & Space(1) & glMID_BACKTRACK_SPACE & Space(1) & "K"
    lblDefHeapSpace = glHEAP_SPACE & Space(1) & "-" & Space(1) & glMID_HEAP_SPACE & Space(1) & "K"
    lblDefInputBuffer = glINPUT_SPACE & Space(1) & "-" & Space(1) & glMID_INPUT_SPACE & Space(1) & "K"
    lblDefLocalSpace = glLOCAL_SPACE & Space(1) & "-" & Space(1) & glMID_LOCAL_SPACE & Space(1) & "K"
    lblDefOutputBuffer = glOUTPUT_SPACE & Space(1) & "-" & Space(1) & glMID_OUTPUT_SPACE & Space(1) & "K"
    lblDefProgramSpace = glPROGRAM_SPACE & Space(1) & "-" & Space(1) & glMID_PROGRAM_SPACE & Space(1) & "K"
    lblDefTextSpace = glTEXT_SPACE & Space(1) & "-" & Space(1) & glMID_TEXT_SPACE & Space(1) & "K"
    
    lblAbMaxBackTrackSpace = glMAX_BACKTRACK_SPACE & Space(1) & "K"
    lblAbMaxHeapSpace = glMAX_HEAP_SPACE & Space(1) & "K"
    lblAbMaxInputBuffer = glMAX_INPUT_SPACE & Space(1) & "K"
    lblAbMaxLocalSpace = glMAX_LOCAL_SPACE & Space(1) & "K"
    lblAbMaxOutputBuffer = glMAX_OUTPUT_SPACE & Space(1) & "K"
    lblAbMaxProgramSpace = glMAX_PROGRAM_SPACE & Space(1) & "K"
    lblAbMaxTextSpace = glMAX_TEXT_SPACE & Space(1) & "K"
    
End Sub

'-----------------------------------------------------------------------------
Public Sub Display(ByVal lTrialId As Long, _
                        ByVal sTrialName As String, _
                        ByVal enAccessMode As eSDAccessMode, _
                        Optional ByVal sMessage As String = "")
'-----------------------------------------------------------------------------
'lTrialID = clinicaltrialid of selected trial
'sTrialName = clinicaltrial name
'sMessage = displays message (if any) when the form is unloaded
' NCJ 13 Jun 06 - Added study access mode
'-----------------------------------------------------------------------------
    
    msMessage = sMessage
    
    'initialise class
    Set moArezzoMemory = New clsAREZZOMemory
    
    'get memory settings for selected study
    ' NCJ 28 Jan 03 - Pass in DB connection string
    Call moArezzoMemory.Load(lTrialId, goUser.CurrentDBConString)
     
    Me.Caption = "AREZZO Memory Settings" & "-" & sTrialName
    
    'load values from database
    DisplayPrologValues enAccessMode
    
    'load and display minimum and maximum values
    DisplayMinMaxValues
    
    mbValuesSaved = True
    mbOKClicked = False
    mbValueChanged = False
    mbAnythingSaved = False
    cmdApply.Enabled = mbAnythingSaved
    
    If enAccessMode < sdReadWrite Then
        ' They can't touch anything
        cmdOK.Enabled = False
        cmdReset.Enabled = False
    End If
    
    Me.Show vbModal

End Sub

'--------------------------------------------------------------------------
Private Sub ResetTextBox(oTextBox As TextBox, sAREZZOMemory As String)
'--------------------------------------------------------------------------
'sets the textbox value to original (as stored in database)
'--------------------------------------------------------------------------
    
    Select Case sAREZZOMemory
        Case gsBACKTRACK_SPACE
            oTextBox.Text = moArezzoMemory.BacktrackSpace
        Case gsHEAP_SPACE
            oTextBox.Text = moArezzoMemory.HeapSpace
        Case gsINPUT_SPACE
            oTextBox.Text = moArezzoMemory.InputSpace
        Case gsLOCAL_SPACE
            oTextBox.Text = moArezzoMemory.LocalSpace
        Case gsOUTPUT_SPACE
             oTextBox.Text = moArezzoMemory.OutputSpace
        Case gsTEXT_SPACE
             oTextBox.Text = moArezzoMemory.TextSpace
        Case gsPROGRAM_SPACE
             oTextBox.Text = moArezzoMemory.ProgramSpace
    End Select

End Sub
