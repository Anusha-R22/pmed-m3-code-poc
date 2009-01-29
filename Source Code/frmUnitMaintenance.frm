VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUnitMaintenance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Units and Conversion Factors"
   ClientHeight    =   4770
   ClientLeft      =   7575
   ClientTop       =   7530
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4770
   ScaleWidth      =   7365
   Begin MSComctlLib.ListView lvwFactor 
      Height          =   2235
      Left            =   3300
      TabIndex        =   24
      Top             =   900
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3942
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdNewFactor 
      Caption         =   "New"
      Height          =   280
      Left            =   6550
      TabIndex        =   22
      Top             =   3200
      Width           =   735
   End
   Begin VB.CommandButton cmdEditFactor 
      Caption         =   "Edit"
      Height          =   280
      Left            =   5760
      TabIndex        =   21
      Top             =   3200
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Insert/Edit Conversion Factors"
      Height          =   1215
      Left            =   3300
      TabIndex        =   12
      Top             =   3480
      Width           =   4000
      Begin VB.CommandButton cmdAddFactor 
         Caption         =   "Add"
         Height          =   280
         Left            =   3120
         TabIndex        =   17
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton cmdCancelFactor 
         Caption         =   "Cancel"
         Height          =   280
         Left            =   2280
         TabIndex        =   16
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtFactor 
         Height          =   285
         Left            =   2280
         MaxLength       =   12
         TabIndex        =   15
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox cboToUnit 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   480
         Width           =   1000
      End
      Begin VB.ComboBox cboFromUnit 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   480
         Width           =   1000
      End
      Begin VB.Label Label6 
         Caption         =   "Conversion factor"
         Height          =   255
         Left            =   2280
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "To Unit"
         Height          =   255
         Left            =   1200
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "From Unit"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraAddEditUnit 
      Caption         =   "Add/Edit Unit"
      Height          =   1215
      Left            =   100
      TabIndex        =   8
      Top             =   3480
      Width           =   3100
      Begin VB.CommandButton cmdAddUnit 
         Caption         =   "Add"
         Height          =   280
         Left            =   2280
         TabIndex        =   11
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton cmdCancelUnit 
         Caption         =   "Cancel"
         Height          =   280
         Left            =   1440
         TabIndex        =   10
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtUnit 
         Height          =   285
         Left            =   120
         MaxLength       =   15
         TabIndex        =   9
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label7 
         Caption         =   "New or changed Unit"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdNewUnit 
      Caption         =   "New"
      Height          =   280
      Left            =   2470
      TabIndex        =   7
      Top             =   3200
      Width           =   735
   End
   Begin VB.CommandButton cmdEditUnit 
      Caption         =   "Edit"
      Height          =   280
      Left            =   1700
      TabIndex        =   6
      Top             =   3200
      Width           =   735
   End
   Begin VB.ListBox lstUnits 
      Height          =   2200
      Left            =   1700
      TabIndex        =   3
      Top             =   900
      Width           =   1500
   End
   Begin VB.ListBox lstUnitClasses 
      Height          =   2200
      Left            =   100
      TabIndex        =   2
      Top             =   900
      Width           =   1500
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Conversion Factors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3300
      TabIndex        =   5
      Top             =   500
      Width           =   4000
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Units"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1700
      TabIndex        =   4
      Top             =   500
      Width           =   1500
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Unit Class"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   100
      TabIndex        =   1
      Top             =   500
      Width           =   1500
   End
   Begin VB.Label lblUnitMaintenance 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   7215
   End
End
Attribute VB_Name = "frmUnitMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998-2007. All Rights Reserved
'   File:       frmUnitMaintenance.frm
'   Author:     Mo Morris, July 1997
'   Purpose:    Maintain units and unit conversion factors
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
 '  1       Mo Morris              18/07/97
'   2       Andrew Newbigging      10/09/97
'   3       Mo Morris              18/09/97
'           Mo Morris               8/10/98     gsMacro_DATABASE replaced by guser.DataBasePath
'   Mo Morris   15/11/99    DAO to ADO conversion
'   Mo Morris   10/12/99    Whole form re-written with out DBGrids and Data controls
'   TA 29/04/2000   subclassing removed
'   TA 20/06/2000 SR3628:   maxlength property of txtUnit set to 15
'                            also sql altered when checking for whether unit name already exists
'   NCJ 13/10/00 - Do not allow changes to units that are in use (in Clinical tests or data items)
'   NCJ 27/10/00 - Do case-insensitive check for existing units in Oracle
'   NCJ 3/11/00 SR4013 - Check for existing conversion factor before adding
'           Also major tidy up of functionality
'   TA 13/12/2000 - Changed conversion factor grid to listview
' ASH 12/2/2003 - Added check for max value allowable for conversion factors
' NCJ 14 Feb 03 - Do not use Val because it doesn't work with Regional Settings
' NCJ 26 Feb 07 - Bug 2774 - Screen out single quotes in Unit name
'------------------------------------------------------------------------------------'

Option Explicit

Private Const m_COL_FROMUNIT = 0
Private Const m_COL_TOUNIT = 1
Private Const m_COL_FACTOR = 2

' Store the current unit class
Private msCurrentUnitClass As String

'---------------------------------------------------------------------
Private Sub cboFromUnit_Click()
'---------------------------------------------------------------------

    Call EnableAddFactorButton

End Sub

'---------------------------------------------------------------------
Private Sub cboToUnit_Click()
'---------------------------------------------------------------------

    Call EnableAddFactorButton

End Sub

'---------------------------------------------------------------------
Private Sub cmdAddFactor_Click()
'---------------------------------------------------------------------
' NCJ 3 Nov 2000 - SR4013 Must check for uniqueness of new conversion factor
'---------------------------------------------------------------------
Dim sSQL As String
Dim sEditFromUnit As String
Dim sEditToUnit As String
Dim dblEditFactor As String
Dim sMsg As String
Dim dblNewFactor As Single
Dim sFactor As String

Const dblMAX_CONVERSION_FACTOR As Double = 999999.9

    On Error GoTo ErrHandler

    sFactor = Trim(txtFactor.Text)
    
    'Validate the Conversion factor
    If Not IsNumeric(sFactor) Then
        sMsg = "The conversion factor must be numeric"
        Call DialogError(sMsg)
        txtFactor.SetFocus
        Exit Sub
    End If
    
    ' NCJ 14 Feb 03 - Cannot use Val function because it does not recognise regional settings! Use CDbl instead
    dblNewFactor = CDbl(sFactor)
    
    'ASH 7/2/2003 Do not allow whole numbers greater than dblMAX_CONVERSION_FACTOR
    If dblNewFactor > dblMAX_CONVERSION_FACTOR Then
        sMsg = "The conversion factor must not be greater than " & dblMAX_CONVERSION_FACTOR
        Call DialogError(sMsg)
        txtFactor.SetFocus
        Exit Sub
    End If
    
    If dblNewFactor <= 0 Then
        sMsg = "The conversion factor must be a positive number"
        Call DialogError(sMsg)
        txtFactor.SetFocus
        Exit Sub
    End If
    
    'Check for FromUnit and ToUnit being the same
    If cboFromUnit.Text = cboToUnit.Text Then
        sMsg = "The From Unit and the To Unit cannot be the same"
        Call DialogError(sMsg)
        Exit Sub
    End If
    
    If cmdAddFactor.Caption = "Change" Then
        ' Editing - check there's something changed
        sEditFromUnit = lvwFactor.SelectedItem.Text
        sEditToUnit = lvwFactor.SelectedItem.SubItems(m_COL_TOUNIT)
        dblEditFactor = CDbl(lvwFactor.SelectedItem.SubItems(m_COL_FACTOR))

        If (cboFromUnit.Text = sEditFromUnit) _
          And (cboToUnit.Text = sEditToUnit) _
          And (dblNewFactor = dblEditFactor) Then
            sMsg = "No changes have been made to the conversion factor that you are editing"
            Call DialogInformation(sMsg)
            txtFactor.SetFocus
            Exit Sub
        End If
        
        'It's an edit, so delete the old value before inserting the new one
        sSQL = "DELETE FROM UnitConversionFactors " _
            & "WHERE FromUnit = '" & sEditFromUnit & "' " _
            & "AND ToUnit = '" & sEditToUnit & "' " _
            & "AND UnitClass = '" & msCurrentUnitClass & "'"
        MacroADODBConnection.Execute sSQL
    End If
    
    ' Check that it doesn't already exist (when Adding or Editing)
    If ConversionFactorExists(cboFromUnit.Text, cboToUnit.Text) Then
        Call DialogInformation("A conversion factor for these units already exists")
        Exit Sub
    End If
    
    ' Everything OK - insert the new/changed conversion factor
    sSQL = "INSERT INTO UnitConversionFactors (FromUnit, ToUnit, ConversionFactor, UnitClass) " _
            & "VALUES ('" & cboFromUnit.Text & "','" & cboToUnit.Text & "'," _
            & ConvertLocalNumToStandard(CStr(dblNewFactor)) & ",'" & msCurrentUnitClass & "')"
    MacroADODBConnection.Execute sSQL
    
    'clear up after the add/change
    Call EnableFactorEditing(False)
    
    RefreshConversionFactors
   
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdAddFactor_Click")
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
Private Function ConversionFactorExists(ByVal sFromUnit As String, _
                                    ByVal sToUnit As String) As Boolean
'---------------------------------------------------------------------
' Returns TRUE if conversion factor already exists for this FromUnit & ToUnit
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bExists As Boolean

    bExists = False
    sSQL = "SELECT * FROM UnitConversionFactors WHERE " _
            & GetSQLStringEquals("FromUnit", sFromUnit) _
            & " AND " & GetSQLStringEquals("ToUnit", sToUnit)
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    If rsTemp.RecordCount <> 0 Then
        bExists = True
    End If
    rsTemp.Close
    Set rsTemp = Nothing
    
    ConversionFactorExists = bExists
    
End Function

'---------------------------------------------------------------------
Private Sub cmdAddUnit_Click()
'---------------------------------------------------------------------
' NCJ 31/5/00 - Use gsCANNOT_CONTAIN_INVALID_CHARS
' NCJ 27/10/00 - Do case-insensitive check for Oracle
' NCJ 26 Feb 07 - New valid unit name checker
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim sUnit As String

    On Error GoTo ErrHandler

    sUnit = Trim(txtUnit.Text)
    
    If sUnit = "" Then
        Call DialogError("Please enter a unit")
        txtUnit.SetFocus
        Exit Sub
    End If
    
    'Validate the unit for characters that will upset SQL
    ' NCJ 26 Feb 07 - Bug 2774 - Use new routine which screens for single quotes too
    If Not ValidUnitName(sUnit) Then
'    If Not gblnValidString(sUnit, valOnlySingleQuotes) Then
'        Call DialogError(sUnit & gsCANNOT_CONTAIN_INVALID_CHARS)
        If cmdAddUnit.Caption = "Change" Then
            txtUnit.Text = lstUnits.Text
        Else
            txtUnit.Text = ""
        End If
        txtUnit.SetFocus
        Exit Sub
    End If
    
    'check that the new/changed unit does not already exist
    ' NCJ 27/10/00 - Case-insensitive for Oracle
    sSQL = "SELECT Unit FROM Units WHERE " & GetSQLStringEquals("Unit", sUnit)
     'TA extra where clause removed as values for unit must be unique
    ' & "AND UnitClass = '" & msCurrentUnitClass & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    If rsTemp.RecordCount <> 0 Then
        Call DialogError("The unit '" & sUnit & "' already exists")
        If cmdAddUnit.Caption = "Change" Then
            txtUnit.Text = lstUnits.Text
        Else
            txtUnit.Text = ""
        End If
        txtUnit.SetFocus
        Exit Sub
    End If
    
    If cmdAddUnit.Caption = "Change" Then
        If sUnit = lstUnits.Text Then
            Call DialogInformation("No changes have been made to the unit that you are editing")
            txtUnit.SetFocus
            Exit Sub
        End If
        'Its an edit, so delete the old value before inserting the new one
        sSQL = "DELETE FROM Units " _
            & "WHERE Unit = '" & lstUnits.Text & "' " _
            & "AND UnitClass = '" & msCurrentUnitClass & "'"
        MacroADODBConnection.Execute sSQL
    End If
    
    'Insert the new/changed value
    sSQL = "INSERT INTO Units (Unit,UnitClass) " _
            & "VALUES ('" & sUnit & "','" & msCurrentUnitClass & "')"
    MacroADODBConnection.Execute sSQL
    
    'clear up after the add/change
    Call EnableUnitEditing(False)
    
'    txtUnit.Text = ""
'    txtUnit.Locked = True
'    cmdCancelUnit.Enabled = False
'    cmdAddUnit.Caption = "Add"
'    cmdAddUnit.Enabled = False
'    lstUnits.Enabled = True    ' NCJ 13/10/00
'    lstUnitClasses.Enabled = True    ' NCJ 13/10/00
    
    RefreshUnits
'    cmdEditUnit.Enabled = False

    'TA 12/12/2000: enable new factor button if appropriate
    Call EnableNewFactorButton
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdAddUnit_Click")
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
Private Function ValidUnitName(sUnit As String) As Boolean
'---------------------------------------------------------------------
' NCJ 26 Feb 07 - Bug 2774 - Disallow unit names with single quotes
' Gives error dialog and returns False if sUnit is invalid
' Returns True if sUnit is OK
'---------------------------------------------------------------------
Dim sMsg As String

    On Error GoTo ErrLabel
    
    sMsg = ""
    ValidUnitName = True
    
    If Not gblnValidString(sUnit, valOnlySingleQuotes) Then
        sMsg = gsCANNOT_CONTAIN_INVALID_CHARS
    ElseIf InStr(sUnit, "'") > 0 Then       ' Extra check for single quotes
        sMsg = " may not contain single quotes."
    End If
    
    If sMsg > "" Then
        DialogError "Unit name" & sMsg
        ValidUnitName = False
    End If

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmUnitMaintenance.ValidUnitName"
    
End Function

'---------------------------------------------------------------------
Private Function UnitIsInUse(ByVal sUnit As String, ByRef sMsg As String) As Boolean
'---------------------------------------------------------------------
' NCJ 13/10/00 - Returns TRUE if unit is in use somewhere
' i.e. in a Data Item, Clinical Test or Conversion Factor
' If in use, sMsg will be relevant message
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bInUse As Boolean

    On Error GoTo ErrHandler
    bInUse = False
    sMsg = ""
    
    ' Check Data Items
    sSQL = "SELECT Count(*) FROM DataItem WHERE " _
            & "UnitOfMeasurement = '" & sUnit & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection
    If rsTemp.Fields(0) > 0 Then
        bInUse = True
        sMsg = sMsg & "question definition"
    End If
    
    ' Check Clinical Tests if not used in data items
    If Not bInUse Then
        sSQL = "SELECT Count(*) FROM ClinicalTest WHERE " _
                & "Unit = '" & sUnit & "'"
        Set rsTemp = New ADODB.Recordset
        rsTemp.Open sSQL, MacroADODBConnection
        If rsTemp.Fields(0) > 0 Then
            bInUse = True
            sMsg = sMsg & "clinical test definition"
        End If
    End If
    
    ' Check Conversion factors if not used in data items or clinical tests
    If Not bInUse Then
        sSQL = "SELECT Count(*) FROM UnitConversionFactors WHERE " _
                & "FromUnit = '" & sUnit & "' OR " _
                & "ToUnit = '" & sUnit & "'"
        Set rsTemp = New ADODB.Recordset
        rsTemp.Open sSQL, MacroADODBConnection
        If rsTemp.Fields(0) > 0 Then
            bInUse = True
            sMsg = sMsg & "conversion factor definition"
        End If
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    If bInUse Then
        sMsg = "The unit '" & sUnit & "' occurs in a " & sMsg & ", " _
                 & vbCrLf & "and may not be changed while it is still in use"
    End If
    
    UnitIsInUse = bInUse
    
Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "UnitIsInUse(" & sUnit & ")")
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
Private Sub cmdCancelFactor_Click()
'---------------------------------------------------------------------

    Call EnableFactorEditing(False)
    Call RefreshConversionFactors ' To reset grid selection

End Sub

'---------------------------------------------------------------------
Private Sub cmdCancelUnit_Click()
'---------------------------------------------------------------------

    Call EnableUnitEditing(False)

End Sub

'---------------------------------------------------------------------
Private Sub cmdEditFactor_Click()
'---------------------------------------------------------------------
' set up factor editing
'---------------------------------------------------------------------
Dim sFromUnit As String
Dim sToUnit As String
Dim i As Long

    Call EnableFactorEditing(True)
    cmdAddFactor.Caption = "Change"

    txtFactor.Text = lvwFactor.SelectedItem.SubItems(m_COL_FACTOR)
    sFromUnit = lvwFactor.SelectedItem.Text
    sToUnit = lvwFactor.SelectedItem.SubItems(m_COL_TOUNIT)
    
    For i = 0 To cboFromUnit.ListCount - 1
        If cboFromUnit.List(i) = sFromUnit Then
            cboFromUnit.ListIndex = i
            Exit For
        End If
    Next
    
    For i = 0 To cboToUnit.ListCount - 1
        If cboToUnit.List(i) = sToUnit Then
            cboToUnit.ListIndex = i
            Exit For
        End If
    Next
    


End Sub

'---------------------------------------------------------------------
Private Sub cmdEditUnit_Click()
'---------------------------------------------------------------------
' NCJ 13 Oct 00 - Don't allow edits on units that are in use
'---------------------------------------------------------------------
Dim sMsg As String

    ' NCJ 13 Oct 00
    If UnitIsInUse(lstUnits.Text, sMsg) Then
        Call DialogError(sMsg, "Unit Maintenance")
    Else
        Call EnableUnitEditing(True)
        cmdAddUnit.Caption = "Change"
        txtUnit.Text = lstUnits.Text
    End If

End Sub

'---------------------------------------------------------------------
Private Sub cmdNewFactor_Click()
'---------------------------------------------------------------------

    Call EnableFactorEditing(True)
    
    cmdAddFactor.Caption = "Add"

End Sub

'---------------------------------------------------------------------
Private Sub cmdNewUnit_Click()
'---------------------------------------------------------------------

    Call EnableUnitEditing(True)
    
    cmdAddUnit.Caption = "Add"


End Sub

'---------------------------------------------------------------------
Private Sub lstUnits_Click()
'---------------------------------------------------------------------
    
    Call EnableEditUnitButton

End Sub
 
'---------------------------------------------------------------------
Private Sub EnableEditUnitButton()
'---------------------------------------------------------------------
    
    If lstUnits.ListIndex > -1 Then
        cmdEditUnit.Enabled = True
    Else
        cmdEditUnit.Enabled = False
    End If

 End Sub

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler
       
    Me.Icon = frmMenu.Icon
    
    lblUnitMaintenance.Caption = " Select a Unit Class before making changes to Units and Conversion Factors"
    
    'populate the UnitClass listbox
    sSQL = "SELECT * from UnitClasses"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsTemp.EOF
        lstUnitClasses.AddItem rsTemp!UnitClass
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing
    
    ' Select first unit class (this will force a refresh of conversion factors)
    lstUnitClasses.ListIndex = 0
    
    ' Initially can't edit anything
    Call EnableUnitEditing(False)
    Call EnableFactorEditing(False)
    
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

'---------------------------------------------------------------------
Private Sub lstUnitClasses_Click()
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
'    ClearTextDisableCommands

    msCurrentUnitClass = lstUnitClasses.Text
    
    'Refresh contents of lstUnits and grdUnitConversionFactors based on selected UnitClass
    RefreshUnits
    RefreshConversionFactors
    Call EnableEditUnitButton
    Call EnableNewFactorButton
    cmdEditFactor.Enabled = False
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "lstUnitClasses_Click")
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
Public Sub RefreshUnits()
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler

    'clear Unit listboxes and combos
    lstUnits.Clear
    cboFromUnit.Clear
    cboToUnit.Clear
    
    sSQL = "SELECT DISTINCT  Unit, UnitClass " _
        & "From Units " _
        & "WHERE UnitClass = '" & msCurrentUnitClass & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Do While Not rsTemp.EOF
        lstUnits.AddItem rsTemp!Unit
        cboFromUnit.AddItem rsTemp!Unit
        cboToUnit.AddItem rsTemp!Unit
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "RefreshUnits")
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
Public Sub RefreshConversionFactors()
'---------------------------------------------------------------------
' refill the conversion factors listview
'---------------------------------------------------------------------
Dim sSQL As String
Dim tblFactors As clsDataTable

    On Error GoTo ErrHandler
      
    sSQL = "SELECT DISTINCT FromUnit, ToUnit, ConversionFactor, UnitClass " _
        & " FROM UnitConversionFactors " _
        & " WHERE UnitClass = '" & msCurrentUnitClass & "'" _
        & " ORDER BY FromUnit, ToUnit"

    Set tblFactors = TableFromSQL(sSQL, RecordBuild("From Unit", "To Unit", "Conversion Factor"))
    Call TableToListView(lvwFactor, tblFactors)
    
    txtFactor.Text = ""
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "RefreshConversionFactors")
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
Private Sub lvwFactor_ItemClick(ByVal Item As MSComctlLib.ListItem)
'---------------------------------------------------------------------

    cmdEditFactor.Enabled = True
    
End Sub

'---------------------------------------------------------------------
Private Sub txtFactor_Change()
'---------------------------------------------------------------------
'---------------------------------------------------------------------

    Call EnableAddFactorButton

End Sub

'---------------------------------------------------------------------
Private Sub EnableAddFactorButton()
'---------------------------------------------------------------------
' Enable the Add/Change Factor button (assume in Edit mode)
'---------------------------------------------------------------------

    If Trim(txtFactor.Text) <> "" _
        And (cboFromUnit.ListIndex <> -1) _
        And (cboToUnit.ListIndex <> -1) Then
            cmdAddFactor.Enabled = True
    Else
            cmdAddFactor.Enabled = False
    End If

End Sub

'---------------------------------------------------------------------
Private Sub txtUnit_Change()
'---------------------------------------------------------------------

    Call EnableAddUnitButton

End Sub

'---------------------------------------------------------------------
Private Sub EnableAddUnitButton()
'---------------------------------------------------------------------
' Enable the Add/Change Unit button (assume in Edit mode)
'---------------------------------------------------------------------
    
    If Trim(txtUnit.Text) <> "" Then
        cmdAddUnit.Enabled = True
    Else
        cmdAddUnit.Enabled = False
    End If

End Sub

'---------------------------------------------------------------------
Private Sub EnableNewFactorButton()
'---------------------------------------------------------------------
' Only allow new conversion factor if at least two units
' (Assume in View mode)
'---------------------------------------------------------------------

        cmdNewFactor.Enabled = (lstUnits.ListCount > 1)

End Sub

'---------------------------------------------------------------------
Private Sub SetViewMode(ByVal bEnable As Boolean)
'---------------------------------------------------------------------
' NCJ 3/11/00 - Enable/Disable "view" mode (i.e. not editing anything)
' NB This is called from EnableUnitEditing and EnableFactorEditing
'---------------------------------------------------------------------
    
    ' These two are always false
    cmdEditUnit.Enabled = False
    cmdEditFactor.Enabled = False
    If bEnable Then
        Call EnableNewFactorButton
    Else
        cmdNewFactor.Enabled = False
    End If
    cmdNewUnit.Enabled = bEnable
    lvwFactor.Enabled = bEnable
    lstUnits.Enabled = bEnable
    lstUnitClasses.Enabled = bEnable
    
End Sub

'---------------------------------------------------------------------
Private Sub EnableUnitEditing(ByVal bEnable As Boolean)
'---------------------------------------------------------------------
' NCJ 3/11/00 - Enable/Disable things when editing/adding units
'---------------------------------------------------------------------

    Call SetViewMode(Not bEnable)
    cmdCancelUnit.Enabled = bEnable
    txtUnit.Enabled = bEnable
    txtUnit.Text = ""
    ' Special ones for entering/exiting edit mode
    If bEnable Then
        txtUnit.SetFocus
        Call EnableAddUnitButton
    Else
        cmdAddUnit.Enabled = False
        Call EnableEditUnitButton
    End If

End Sub

'---------------------------------------------------------------------
Private Sub EnableFactorEditing(ByVal bEnable As Boolean)
'---------------------------------------------------------------------
' NCJ 3/11/00 - Enable/Disable things when editing/adding conversion factors
'---------------------------------------------------------------------

    Call SetViewMode(Not bEnable)
    cmdCancelFactor.Enabled = bEnable
    cboFromUnit.ListIndex = -1
    cboToUnit.ListIndex = -1
    cboFromUnit.Enabled = bEnable
    cboToUnit.Enabled = bEnable
    txtFactor.Enabled = bEnable
    txtFactor.Text = ""
    If bEnable Then
        Call EnableAddFactorButton
        txtFactor.SetFocus
    Else
        cmdAddFactor.Enabled = False
    End If
    
End Sub



