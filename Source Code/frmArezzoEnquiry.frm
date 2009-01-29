VERSION 5.00
Begin VB.Form frmArezzoEnquiry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enquiry"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10425
   ControlBox      =   0   'False
   Icon            =   "frmArezzoEnquiry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   420
      TabIndex        =   13
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Frame fraCategory 
      Caption         =   "Frame1"
      Height          =   2175
      Left            =   5160
      TabIndex        =   8
      Top             =   2640
      Width           =   4695
      Begin VB.ListBox lstValues 
         Height          =   1620
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   4455
      End
   End
   Begin VB.Frame fraBoolean 
      Caption         =   "Frame1"
      Height          =   1815
      Left            =   5160
      TabIndex        =   5
      Top             =   360
      Width           =   4695
      Begin VB.OptionButton optFalse 
         Caption         =   "False"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1080
         Width           =   3975
      End
      Begin VB.OptionButton optTrue 
         Caption         =   "True"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   3975
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter value"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Frame fraEdit 
      Caption         =   "Frame1"
      Height          =   2175
      Left            =   360
      TabIndex        =   1
      Top             =   3240
      Width           =   4695
      Begin VB.TextBox txtValue 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   960
         Width           =   4455
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Data Unit"
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label lblEdit 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter a value:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   3255
      End
   End
   Begin VB.ListBox lstDItems 
      Height          =   1620
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   4695
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Task description"
      Height          =   855
      Left            =   360
      TabIndex        =   11
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Values have been requested for the following data items:"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   1080
      Width           =   4695
   End
End
Attribute VB_Name = "frmArezzoEnquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------
' File: frmArezzoEnquiry.frm
' Copyright InferMed Ltd 1999-2003 All Rights Reserved
' Author: Nicky Johns, InferMed
' Purpose: Deal with Arezzo enquiries & decision data in MACRO Data Management
'-----------------------------------------
' REVISIONS
'   NCJ 28-29 Sep 99 - Initial Development
'   WillC  10/11/99 Added errhandlers
' MACRO 2.2
'   NCJ 27 Sep 01 - Changed to use new goArezzo in MACRO 2.2
'   NCJ 1 Oct 01 - Changed RefreshMe to Display
'   NCJ 12 Oct 01 - Return TRUE if there was no data requested
'   NCJ 17 Jan 02 - Added Hourglass suspend/resume (Buglist 2.2.3, Bug 25)
'   NCJ 31 Jan 03 - Added Cancel button, and use oArezzo passed in
'   NCJ 10 Feb 03 - Use new DEBS routine to add data
'-----------------------------------------

Option Explicit

Private moEnquiry As TaskInstance
' Type of data entry
Private msWidgetType As String
Private mbOKClicked As Boolean

' NCJ 31 Jan 03
Private moArezzo As Arezzo_DM

'-----------------------------------------
Public Function Display(oTask As TaskInstance, oArezzo As Arezzo_DM) As Boolean
'-----------------------------------------
' Refresh and display for the given task (decision or enquiry)
' Return TRUE if data was entered, or there was no data,
' or FALSE if there was data requested but none was entered
'-----------------------------------------

    On Error GoTo ErrHandler
    
    Set moArezzo = oArezzo
    
    mbOKClicked = False
    
    Set moEnquiry = oTask
    If oTask.TaskType = "enquiry" Then
        Me.Caption = "Enquiry - " & oTask.Name
    Else
        Me.Caption = "Decision - " & oTask.Name
    End If
    lblDesc.Caption = oTask.Description
    
    ' Update the list box
    If FillListbox Then
        ' There were some data items
        ' NCJ 17 Jan 02 - Suspend/resume hourglass
        HourglassSuspend
        Me.Show vbModal
        HourglassResume
        
        Display = mbOKClicked
    Else
        ' Default to TRUE (in case of decisions)
        Display = True
    End If
    
    
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
            End
    End Select
    
End Function

'---------------------------------------------------------
Private Function FillListbox() As Boolean
'---------------------------------------------------------
' Fill the list box with the data requested for this task
' Returns TRUE if there is requested data
' Returns FALSE if no data
'---------------------------------------------------------
Dim vDItem As Variant

    On Error GoTo ErrHandler
    
    ' Hide the data entry areas
    fraBoolean.Visible = False
    fraCategory.Visible = False
    fraEdit.Visible = False
    cmdEnter.Enabled = False
    
    lstDItems.Clear
    If moEnquiry.RequestedData.Count > 0 Then
        For Each vDItem In moEnquiry.RequestedData
            lstDItems.AddItem CStr(vDItem)
        Next vDItem
        ' Select first item
        lstDItems.ListIndex = 0
        FillListbox = True
    Else        ' No requested data
        FillListbox = False
    End If
    
Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "FillListbox")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Function

'---------------------------------------------------------
Private Sub cmdCancel_Click()
'---------------------------------------------------------
' They decided to Cancel
' NB They might have already entered some data
'---------------------------------------------------------

    Unload Me

End Sub

'---------------------------------------------------------
Private Sub cmdEnter_Click()
'---------------------------------------------------------
' Enter value for the currently selected data item
'---------------------------------------------------------
Dim sValue As String
Dim sDItem As String

    On Error GoTo ErrHandler
    
    sDItem = GetSelectedItem(lstDItems)
'    With moArezzo.ALM.GuidelineInstance
    
    Select Case msWidgetType
    Case "boolean"
        If optTrue.Value = True Then
            sValue = optTrue.Caption
        Else
            sValue = optFalse.Caption
        End If
    Case "edit"
        sValue = txtValue.Text
    Case "category"
        sValue = GetSelectedItem(lstValues)
    End Select
        
'        Call .colDataValues.Add(sDItem, QuoteIfNecessary(sValue), False)
    Call moArezzo.AddNonMacroData(sDItem, QuoteIfNecessary(sValue))

'    End With
    
    ' Check for errors
    If DataIsOK(sDItem) Then
        ' Store that we did something
        mbOKClicked = True
        ' Refresh the listbox - if empty, we go away
        If FillListbox Then
            ' There's still something to do
        Else
            ' Nothing left to do
            Unload Me
        End If
    Else
        ' There was an error
        ' If an edit field, select the errant value
        If msWidgetType = "edit" Then
            txtValue.SelStart = 0
            txtValue.SelLength = Len(txtValue.Text)
            txtValue.SetFocus
        End If
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdEnter_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
' Put the frames where they're supposed to be
'---------------------------------------------------------
Const nTop = 3240
Const nLeft = 360

    On Error GoTo ErrHandler
    
    Me.BackColor = glFormColour

    fraBoolean.Top = nTop
    fraBoolean.Left = nLeft
    fraCategory.Top = nTop
    fraCategory.Left = nLeft
    fraEdit.Top = nTop
    fraEdit.Left = nLeft
    
    ' Set correct width
    Me.Width = 5550
    
    ' Centre us on the screen
    Me.Top = (Screen.Height - Me.Height) \ 2
    Me.Left = (Screen.Width - Me.Width) \ 2
    
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
            End
    End Select
    
End Sub

'---------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'---------------------------------------------------------
' Tidy up
'---------------------------------------------------------
    
    Set moEnquiry = Nothing
    Set moArezzo = Nothing

End Sub

'---------------------------------------------------------
Private Sub lstDItems_Click()
'---------------------------------------------------------
' Click on a requested item
' Show an appropriate data entry frame
'---------------------------------------------------------
Dim sDItem As String
Dim sDType As String
Dim sDefaultValue As String
Dim vValue As Variant

    On Error GoTo ErrHandler
    
    sDItem = GetSelectedItem(lstDItems)
    With moArezzo.ALM.GuidelineInstance
    
        sDType = .DataDefinition.DataType(sDItem)
        sDefaultValue = .DataDefinition.DefaultValue(sDItem)
        
        If sDType = "boolean" Then
            'Display option buttons
            msWidgetType = "boolean"
            ' Set up the option button captions
            optTrue.Caption = .DataDefinition.TrueValue(sDItem)
            optFalse.Caption = .DataDefinition.FalseValue(sDItem)
            optTrue.Value = False
            optFalse.Value = False
            ' Show the "boolean" frame
            fraBoolean.Visible = True
            fraEdit.Visible = False
            fraCategory.Visible = False
            fraBoolean.Caption = .DataDefinition.Caption(sDItem)
            ' Select default value (if any)
            If sDefaultValue = optTrue.Caption Then
                optTrue.Value = True
                cmdEnter.Enabled = True
            ElseIf sDefaultValue = optFalse.Caption Then
                optFalse.Value = True
                cmdEnter.Enabled = True
            Else
                ' Leave both unselected
                cmdEnter.Enabled = False
            End If
            
        ElseIf .DataDefinition.RangeValues(sDItem).Count > 0 Then
            ' Display as category
            msWidgetType = "category"
            ' Show the "category" frame
            fraBoolean.Visible = False
            fraEdit.Visible = False
            fraCategory.Visible = True
            fraCategory.Caption = .DataDefinition.Caption(sDItem)
            cmdEnter.Enabled = False
            ' Fill in the available values
            lstValues.Clear
            For Each vValue In .DataDefinition.RangeValues(sDItem)
                lstValues.AddItem CStr(vValue)
                ' Select if it's the default
                If CStr(vValue) = sDefaultValue Then
                    lstValues.ListIndex = lstValues.NewIndex
                    cmdEnter.Enabled = True
                End If
            Next
            
        Else
            'Display an edit field
            msWidgetType = "edit"
            ' Show the "edit" frame
            fraBoolean.Visible = False
            fraEdit.Visible = True
            fraCategory.Visible = False
            fraEdit.Caption = .DataDefinition.Caption(sDItem)
            ' Display default value
            txtValue.Text = sDefaultValue
            ' Ensure initial correct enabling of cmdEnter
            If sDefaultValue > "" Then
                cmdEnter.Enabled = True
            Else
                cmdEnter.Enabled = False
            End If
            lblUnit.Caption = .DataDefinition.Unit(sDItem)
            ' Show data type expected
            Select Case .DataDefinition.DataType(sDItem)
            Case "integer"
                lblEdit.Caption = "Please enter a value (as an integer):"
            Case "real"
                lblEdit.Caption = "Please enter a value (as a number):"
            Case "date", "datetime"
                lblEdit.Caption = "Please enter a value (as a date):"
            Case "time"
                lblEdit.Caption = "Please enter a value (as a time):"
            Case Else
                lblEdit.Caption = "Please enter a value:"
            End Select
            
        End If
    
    End With
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "lstDItems_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Sub
 
'---------------------------------------------------------
Private Function GetSelectedItem(oListBox As ListBox) As String
'---------------------------------------------------------
' Get the currently selected data item from the listbox
'---------------------------------------------------------
Dim nIndex As Integer

    On Error GoTo ErrHandler
    
    nIndex = oListBox.ListIndex
    If nIndex > -1 Then
        GetSelectedItem = oListBox.List(nIndex)
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
            End
    End Select
    
End Function

'---------------------------------------------------------
Private Sub lstValues_Click()
'---------------------------------------------------------
' Select an item from the Values list
'---------------------------------------------------------

    On Error GoTo ErrHandler
    
    If lstValues.ListIndex > -1 Then    ' anything selected?
        cmdEnter.Enabled = True
    Else
        cmdEnter.Enabled = False
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "lstValues_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Sub

'---------------------------------------------------------
Private Sub optFalse_Click()
'---------------------------------------------------------
' Click on "False" option button
'---------------------------------------------------------

    Call CheckOptionSelection
    
End Sub

'---------------------------------------------------------
Private Sub optTrue_Click()
'---------------------------------------------------------
' Click on "True" option button
'---------------------------------------------------------

    Call CheckOptionSelection

End Sub

'---------------------------------------------------------
Private Sub CheckOptionSelection()
'---------------------------------------------------------
' Enable Enter button if either option button is selected
'---------------------------------------------------------

    On Error GoTo ErrHandler
    
    ' See if anything's selected
    If optFalse.Value = True Or optTrue.Value = True Then
        cmdEnter.Enabled = True
    Else
        cmdEnter.Enabled = False
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "CheckOptionSelection")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Sub

'---------------------------------------------------------
Private Sub txtValue_Change()
'---------------------------------------------------------
' They've typed in the value edit field
'---------------------------------------------------------

    On Error GoTo ErrHandler
    
    If Len(Trim(txtValue.Text)) > 0 Then
        cmdEnter.Enabled = True
    Else
        cmdEnter.Enabled = False
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtValue_Change")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Sub

'----------------------------------
Private Function DataIsOK(sDItem As String) As Boolean
'----------------------------------
' Handle data entry errors
' Don't bother reporting warnings here
' Return TRUE if data was OK
' Return FALSE if data was not added because of an error
'----------------------------------
Dim sValue As String
Dim sFlag As String
Dim sMsg As String

    On Error GoTo ErrHandler
    
    Select Case moArezzo.ALM.GuidelineInstance.colDataValues.DataError(sDItem, sValue, sFlag)
    Case -1     ' Type error
        sMsg = "Value " & sDItem & " = " & sValue _
                & vbCrLf & " has not been accepted because it is of the wrong type"
    Case -2     ' Range error (shouldn't happen)
        sMsg = "Not an acceptable value for " & sDItem
    Case -3     ' Validation error
        sMsg = "Value " & sDItem & " = " & sValue _
                & vbCrLf & " has not been accepted because it does not satisfy this condition: " & vbCrLf & vbCrLf _
                & moArezzo.ALM.GuidelineInstance.DataDefinition.MandValidation(sDItem)
    Case Else
        sMsg = ""
    End Select
    
    ' Display error message if any
    If sMsg > "" Then
        DialogInformation sMsg, "AREZZO Data Error"
        DataIsOK = False
    Else
        DataIsOK = True
    End If
    
Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "DataIsOK")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Function

Private Function QuoteIfNecessary(sValue As String) As String
' Quote a data value if it contains spaces
' and it's not already quoted
Dim iSepPos As Integer
    
On Error GoTo ErrHandler
    
    iSepPos = InStr(1, sValue, "'", vbBinaryCompare)
    If iSepPos = 0 Then     ' No quotes already
        iSepPos = InStr(1, sValue, " ", vbBinaryCompare)
        If iSepPos > 0 Then   ' Space was found
            QuoteIfNecessary = "'" & sValue & "'"
        Else
            QuoteIfNecessary = sValue
        End If
    Else
        QuoteIfNecessary = sValue
    End If
    
Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "QuoteIfNecessary")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Function

