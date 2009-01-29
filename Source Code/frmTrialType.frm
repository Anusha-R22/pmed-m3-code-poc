VERSION 5.00
Begin VB.Form frmTrialType 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Study Types"
   ClientHeight    =   4200
   ClientLeft      =   9435
   ClientTop       =   5910
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraStudyType 
      Caption         =   "Existing study types"
      Height          =   2355
      Left            =   60
      TabIndex        =   6
      Top             =   0
      Width           =   5775
      Begin VB.ListBox lstTrialValues 
         Height          =   2010
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   4440
         TabIndex        =   7
         Top             =   1380
         Width           =   1215
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "&New"
         Height          =   375
         Left            =   4440
         TabIndex        =   9
         Top             =   1860
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4500
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Frame fraEdit 
      Caption         =   "Insert/Edit a study type"
      Height          =   1215
      Left            =   60
      TabIndex        =   0
      Top             =   2400
      Width           =   5775
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4440
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "Add"
         Height          =   375
         Left            =   4440
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtEditName 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   300
         Width           =   2895
      End
      Begin VB.Label lblEditName 
         Caption         =   "Study type name:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmTrialType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999. All Rights Reserved
'   File:       frmTrialList.frm
'   Author:     Will Casey August 1999
'   Purpose:    Maintains the table of trialtypes in the Library manager
'               and the table of TrialPhases
'----------------------------------------------------------------------------------------'
' FORM: frmTrialType.frm
' Allows the maintenance of the TrialType table or TrialPhase table
' This form is available from the Parameters menu on frmMenu form.
' Note that the FormType property determines whether it's Types or Phases
'-----------------------------------------------------------
' Revisions:
'   PN  20/09/99    Added moFormIdleWatch object to handle system idle timer resets
'  WillC 11/10/99   Added the error handlers
'   Mo Morris   15/11/99    DAO to ADO conversion
' NCJ 8 Dec 99 - Tidying up (SRs 2305 - 2309)
'               Allow editing of either Trial Types OR Trial Phases
'   Mo Morris   20/12/99    Changes from word 'TRIAL' to 'STUDY'
'   TA 29/04/2000   subclassing removed
'---------------------------------------------------------------------------------------'

Option Explicit
Option Base 0
Option Compare Binary

Private Const msFormTrialType = "Types"
Private Const msFormTrialPhase = "Phases"

' This is "Types" for Trial Types
' or "Phases" for trial phases
Private msFormType As String

' Store the trial type name being edited
' This is empty if we're in "New" mode
Private msValueBeingEdited As String
' Store whether we're in Edit mode
Private mbEditingName As Boolean

'---------------------------------------------------------------------
Public Property Let FormType(sType As String)
'---------------------------------------------------------------------
' The type of the form
' sType should be "Type" or "Phase"
'---------------------------------------------------------------------

    msFormType = sType
    If msFormType = msFormTrialType Then
        Me.Caption = "Study Types"
        fraStudyType.Caption = "Existing study types"
        lblEditName.Caption = "Study type name: "
    Else
        Me.Caption = "Study Phases"
        fraStudyType.Caption = "Existing study phases"
        lblEditName.Caption = "Study phase name: "
    End If
    
End Property

'---------------------------------------------------------------------
Private Function InsertNewValue(sName As String) As Boolean
'---------------------------------------------------------------------
' Insert a new value
' Assume sName is validated
'---------------------------------------------------------------------

    If msFormType = msFormTrialType Then
        ' Trial Type
        InsertNewValue = InsertNewTrial(sName)
    Else
        ' Trial Phase
        InsertNewValue = InsertNewPhase(sName)
    End If
    
End Function

'---------------------------------------------------------------------
Private Function ChangeExistingValue(sName As String, _
                                    sNewName As String) As Boolean
'---------------------------------------------------------------------
' Change existing value
' Assume sNewName is validated
'---------------------------------------------------------------------

    If msFormType = msFormTrialType Then
        ' Trial Type
        ChangeExistingValue = ChangeTrialTypeName(sName, sNewName)
    Else
        ' Trial Phase
        ChangeExistingValue = ChangePhaseName(sName, sNewName)
    End If

End Function

'---------------------------------------------------------------------
Private Sub cmdApply_Click()
'---------------------------------------------------------------------
' They want to apply their edits
'---------------------------------------------------------------------
Dim sName As String
Dim bAppliedOK As Boolean

    On Error GoTo ErrHandler
    
    sName = Trim(txtEditName.Text)    ' Assume already validated
    
    If msValueBeingEdited = "" Then
        ' We're inserting a new one
        bAppliedOK = InsertNewValue(sName)
    Else
        ' We're changing an existing one
        bAppliedOK = ChangeExistingValue(msValueBeingEdited, sName)
    End If
    
    ' Did it happen OK?
    If bAppliedOK Then
        RefreshList
        Call DisableEditing
    Else
        ' Put them back into the edit field
        txtEditName.SetFocus
    End If

    Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdApply_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub cmdCancel_Click()
'---------------------------------------------------------------------
' User is cancelling the current edit
'---------------------------------------------------------------------

    Call DisableEditing
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdEdit_Click()
'---------------------------------------------------------------------
' User has chosen to edit a trial type
' Assume there is a trial type selected in list box
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    msValueBeingEdited = lstTrialValues.List(lstTrialValues.ListIndex)
    If msFormType = msFormTrialType Then
        fraEdit.Caption = " Editing Study Type: " & msValueBeingEdited
    Else
        fraEdit.Caption = " Editing Study Phase: " & msValueBeingEdited
    End If
    Call EnableEditing(msValueBeingEdited)
    cmdApply.Caption = "Change"
    
    Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdEdit_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub cmdInsert_Click()
'---------------------------------------------------------------------
' User wants to insert a new trial type
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    msValueBeingEdited = ""
    If msFormType = msFormTrialType Then
        fraEdit.Caption = " New Study Type "
    Else
        fraEdit.Caption = " New Study Phase "
    End If
    ' Initialise with empty string
    Call EnableEditing("")
    cmdApply.Caption = "Add"
    
    Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdInsert_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub DisableEditing()
'---------------------------------------------------------------------
' Disable editing until they choose Edit or New
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    mbEditingName = False
    msValueBeingEdited = ""
    If msFormType = msFormTrialType Then
        fraEdit.Caption = " Study Type "
    Else
        fraEdit.Caption = " Study Phase "
    End If
    txtEditName.Text = ""
'    txtEditName.Locked = True
    txtEditName.Enabled = False
    ' Only enable Edit if a trial type is selected
    If lstTrialValues.ListIndex > -1 Then
        cmdEdit.Enabled = True
    Else
        cmdEdit.Enabled = False
    End If
    ' Let them enter a new one
    cmdInsert.Enabled = True
    cmdCancel.Enabled = False
    cmdApply.Enabled = False
    
    Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "DisableEditing")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub EnableEditing(sText As String)
'---------------------------------------------------------------------
' Enable editing, prefilling trial type name field with given text
'---------------------------------------------------------------------
   
    On Error GoTo ErrHandler
    
    mbEditingName = True
'    txtEditName.Locked = False
    txtEditName.Enabled = True
    txtEditName.Text = sText
    ' Disable these buttons until we've finished editing
    cmdEdit.Enabled = False
    cmdInsert.Enabled = False
    ' Allow Cancel (cmdApply is set when setting txtEditName.Text)
    cmdCancel.Enabled = True
    txtEditName.SetFocus
    
    Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "EnableEditing")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select

End Sub

Private Sub datTrialType_Validate(Action As Integer, Save As Integer)

End Sub



'---------------------------------------------------------------------
Private Sub lstTrialValues_Click()
'---------------------------------------------------------------------
' Select a trial type (or trial phase)
' If we're not currently editing a name, enable the Edit button
'---------------------------------------------------------------------

    If lstTrialValues.ListIndex > -1 And Not mbEditingName Then
        cmdEdit.Enabled = True
    End If
    
End Sub


'----------------------------------------------------------------------------------------'
Private Function ChangePhaseName(sTrialPhaseName As String, _
                        sNewTrialPhaseName As String) As Boolean
'----------------------------------------------------------------------------------------'
' Change trial phase name
' Assume sNewTrialPhaseName is valid
' Check for uniqueness of new name
' then update database
' Return TRUE if change was successful
'----------------------------------------------------------------------------------------'
    
Dim sSQL As String
Dim iTrialPhaseId As Integer
Dim rsTemp As ADODB.Recordset
Dim rsTrialPhases As ADODB.Recordset

    On Error GoTo ErrHandler
    
     Set rsTrialPhases = New ADODB.Recordset
     Set rsTrialPhases = GetTrialPhases
     
    If DoesNameExistInRecordset(rsTrialPhases, UCase(sNewTrialPhaseName)) = False Then
      ' Open the Recordset to get the necessary record to edit
        sSQL = " SELECT PhaseId, PhaseName " _
                & "FROM TrialPhase " _
                & "WHERE PhaseName = '" & ReplaceQuotes(sTrialPhaseName) & "'"
        Set rsTemp = New ADODB.Recordset
        rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
            
        ' Set the Record Id
        iTrialPhaseId = rsTemp!PhaseId
        
        rsTemp.Close
        Set rsTemp = Nothing
        
        sSQL = "UPDATE TrialPhase SET " _
                & " PhaseName = '" & ReplaceQuotes(sNewTrialPhaseName) _
                & "' WHERE PhaseId = " & iTrialPhaseId
    
        MacroADODBConnection.Execute sSQL
        ChangePhaseName = True
    Else
        MsgBox "A phase of this name exists already", vbOKOnly, "Study Phases"
        ChangePhaseName = False
    End If
    
    rsTrialPhases.Close
    Set rsTrialPhases = Nothing
    
Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "ChangePhaseName")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    

End Function

'----------------------------------------------------------------------------------------'
Private Function ChangeTrialTypeName(sTrialTypeName As String, sNewTrialTypeName As String) As Boolean
'----------------------------------------------------------------------------------------'
' Change trial type name
' Assume sNewTrialTypeName is valid
' Check for uniqueness of new name
' then update database
' Return TRUE if change was successful
'----------------------------------------------------------------------------------------'
    
Dim sSQL As String
Dim iTrialTypeId As Integer
Dim rsTemp As ADODB.Recordset
Dim rsTrialTypes As ADODB.Recordset

    On Error GoTo ErrHandler
    
     Set rsTrialTypes = New ADODB.Recordset
     Set rsTrialTypes = GetTrialTypes
     
    If DoesNameExistInRecordset(rsTrialTypes, UCase(sNewTrialTypeName)) = False Then
      ' Open the Recordset to get the necessary record to edit
        sSQL = " SELECT   TrialTypeId , TrialTypeName " _
            & "FROM TrialType WHERE TrialTypeName = '" & ReplaceQuotes(sTrialTypeName) & "'"
        Set rsTemp = New ADODB.Recordset
        rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
            
        ' Set the Record Id to the Id chosen by the user and the new name to that of the
        ' edited string.
        iTrialTypeId = rsTemp!TrialTypeId
        
        rsTemp.Close
        Set rsTemp = Nothing
        
        sSQL = "UPDATE TrialType SET " _
                & " TrialTypeName = '" & ReplaceQuotes(sNewTrialTypeName) _
                & "' WHERE TrialTypeId = " & iTrialTypeId
    
        MacroADODBConnection.Execute sSQL
        ChangeTrialTypeName = True
    Else
        MsgBox " A study type of this name exists already", vbOKOnly, "Study Types"
        ChangeTrialTypeName = False
    End If
    
    rsTrialTypes.Close
    Set rsTrialTypes = Nothing
    
Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "ChangeTrialTypeName")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Function

'----------------------------------------------------------------------------------------'
Private Sub cmdExit_Click()
'----------------------------------------------------------------------------------------'

    Unload Me
 
End Sub

'----------------------------------------------------------------------------------------'
Private Function InsertNewPhase(sPhaseName As String)
'----------------------------------------------------------------------------------------'
' Insert new trial phase
' Check for uniqueness then add to database
' Assume sPhaseName is valid
' Return TRUE if insert succeeded
'----------------------------------------------------------------------------------------'
Dim rsTrialPhases As ADODB.Recordset

    On Error GoTo ErrHandler
     
    Set rsTrialPhases = New ADODB.Recordset
    Set rsTrialPhases = GetTrialPhases
   
    If DoesNameExistInRecordset(rsTrialPhases, sPhaseName) = False Then
        Call InsertPhaseName(sPhaseName)
        InsertNewPhase = True
    Else
        MsgBox " This study phase exists already", vbOKOnly, "Study Phases"
        InsertNewPhase = False
    End If
    
    rsTrialPhases.Close
    Set rsTrialPhases = Nothing
    
Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdInsert_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select

End Function

'----------------------------------------------------------------------------------------'
Private Function InsertNewTrial(sTrialTypeName As String) As Boolean
'----------------------------------------------------------------------------------------'
' Insert new trial type
' Check for uniqueness then add to database
' Assume sTrialTypeName is valid
' Return TRUE if insert succeeded
'----------------------------------------------------------------------------------------'
Dim rsTrialTypes As ADODB.Recordset

    On Error GoTo ErrHandler
     
    Set rsTrialTypes = New ADODB.Recordset
    Set rsTrialTypes = GetTrialTypes
   
    If DoesNameExistInRecordset(rsTrialTypes, sTrialTypeName) = False Then
        Call InsertTrialTypeName(sTrialTypeName)
        InsertNewTrial = True
    Else
        MsgBox " This study type exists already", vbOKOnly, "Study Types"
        InsertNewTrial = False
    End If
    
    rsTrialTypes.Close
    Set rsTrialTypes = Nothing
    
Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdInsert_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Function

'----------------------------------------------------------------------------------------'
Private Sub InsertTrialTypeName(sTrialTypeName As String)
'----------------------------------------------------------------------------------------'
' Select the greatest trialtypeId add one to it then go ahead with the insert
'----------------------------------------------------------------------------------------'
     
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim iTrialTypeId As Integer

    On Error GoTo ErrHandler

    sSQL = " SELECT (MAX(TrialTypeId) + 1) as NewTrialTypeId " _
        & "FROM TrialType"
        
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If IsNull(rsTemp!NewTrialTypeId) Then    'if no records exist
        iTrialTypeId = 1
    Else
        iTrialTypeId = rsTemp!NewTrialTypeId
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

    'Insert the data into the TrialType table
    sSQL = " INSERT INTO  TrialType" _
            & "(TrialTypeId,TrialTypeName)" _
            & " VALUES(" & iTrialTypeId & ",'" & ReplaceQuotes(sTrialTypeName) & "')"
            
    MacroADODBConnection.Execute sSQL
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "InsertTrialTypeName")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub InsertPhaseName(sPhaseName As String)
'----------------------------------------------------------------------------------------'
' Insert the given new name into the TrialPhase table
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim iPhaseId As Integer

    On Error GoTo ErrHandler

    sSQL = " SELECT (MAX(PhaseId) + 1) as NewPhaseId " _
        & "FROM TrialPhase"
        
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If IsNull(rsTemp!NewPhaseId) Then    'if no records exist
        iPhaseId = 1
    Else
        iPhaseId = rsTemp!NewPhaseId
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

    'Insert the data into the Phase table
    sSQL = " INSERT INTO  TrialPhase" _
            & "(PhaseId,PhaseName)" _
            & " VALUES(" & iPhaseId & ",'" & ReplaceQuotes(sPhaseName) & "')"
            
    MacroADODBConnection.Execute sSQL
    
    Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "InsertPhaseName")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Form_Load()
'----------------------------------------------------------------------------------------'
' Disable buttons, load the listbox
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler
        
    Me.Icon = frmMenu.Icon
    
    mbEditingName = False
    
    Call RefreshList
    
    ' Set up for not editing anything to begin with
    Call DisableEditing
        
    FormCentre Me
   
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


'----------------------------------------------------------------------------------------'
Private Sub RefreshList()
'----------------------------------------------------------------------------------------'
' Read the current valuse into the list box
'----------------------------------------------------------------------------------------'
     
Dim rsTrialValues As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    lstTrialValues.Clear
    
    Set rsTrialValues = New ADODB.Recordset
    If msFormType = msFormTrialType Then
        Set rsTrialValues = GetTrialTypes
    Else
        Set rsTrialValues = GetTrialPhases
    End If
    
    With rsTrialValues
           
      Do Until .EOF = True
        If msFormType = msFormTrialType Then
            lstTrialValues.AddItem rsTrialValues!TrialTypeName
        Else
            lstTrialValues.AddItem rsTrialValues!PhaseName
        End If
        .MoveNext
      Loop
    
    End With
    
    rsTrialValues.Close
    Set rsTrialValues = Nothing
    
    If lstTrialValues.ListCount > 0 Then
        ' Preselect first one
        lstTrialValues.ListIndex = 0
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "RefreshList")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub txtEditName_Change()
'----------------------------------------------------------------------------------------'
' Validate as trial type name
'----------------------------------------------------------------------------------------'
Dim sTemp As String

    On Error GoTo ErrHandler
    
    sTemp = Trim(txtEditName.Text)
    
    If sTemp = "" Then
        cmdApply.Enabled = False
        txtEditName.Tag = ""
    ElseIf IsValidName(sTemp) Then
        cmdApply.Enabled = True
        ' Store as valid value
        txtEditName.Tag = sTemp
    Else
        cmdApply.Enabled = False
        ' Restore previous known valid value
        txtEditName.Text = txtEditName.Tag
        txtEditName.SelStart = Len(txtEditName.Text)
        txtEditName.SelText = ""
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtEditName_Change")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Sub

'----------------------------------------------------------------------------------------'
Private Function GetTrialTypes() As ADODB.Recordset
'----------------------------------------------------------------------------------------'
' Select all the trialtypes from the database
'----------------------------------------------------------------------------------------'
Dim sSQL As String

    On Error GoTo ErrHandler
    
    Set GetTrialTypes = New ADODB.Recordset
    sSQL = "SELECT TrialTypeName from TrialType"
    GetTrialTypes.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

    Exit Function
    
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "GetTrialTypes")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Function

'----------------------------------------------------------------------------------------'
Private Function GetTrialPhases() As ADODB.Recordset
'----------------------------------------------------------------------------------------'
' Select all the trial phases from the database
'----------------------------------------------------------------------------------------'
Dim sSQL As String

    On Error GoTo ErrHandler
    
    Set GetTrialPhases = New ADODB.Recordset
    sSQL = "SELECT PhaseName from TrialPhase"
    GetTrialPhases.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

    Exit Function
    
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "GetTrialPhases")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Function

'----------------------------------------------------------------------------------------'
Private Function IsValidName(sName As String) As Boolean
'----------------------------------------------------------------------------------------'
' Return TRUE if text is valid name for TrialType or TrialPhase
' Displays any necessary messages
'----------------------------------------------------------------------------------------'
Dim sNames As String

    On Error GoTo ErrHandler
    
    IsValidName = False
    
    If msFormType = msFormTrialType Then
        sNames = "Study type names"
    Else
        sNames = "Study phase names"
    End If
    
    If sName > "" Then
        If Not gblnValidString(sName, valOnlySingleQuotes) Then
            MsgBox sNames & gsCANNOT_CONTAIN_INVALID_CHARS, _
                    vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
        ElseIf Not gblnValidString(sName, valAlpha + valNumeric + valSpace) Then
            MsgBox sNames & " may only contain alphanumeric characters", _
                    vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
        ElseIf Len(sName) > 50 Then
            MsgBox sNames & " may not be more than 50 characters", _
                    vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
        Else
            IsValidName = True
        End If
    End If
    
    Exit Function
    
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "IsValidName")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Function

'----------------------------------------------------------------------------------------'
Private Sub txtEditName_KeyPress(KeyAscii As Integer)
'----------------------------------------------------------------------------------------'
' Interpret RETURN to mean cmdApply_Click
'----------------------------------------------------------------------------------------'

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call cmdApply_Click
    End If

End Sub
