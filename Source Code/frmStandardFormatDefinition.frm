VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmStandardFormatDefinition 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Standard Data Formats"
   ClientHeight    =   4230
   ClientLeft      =   4395
   ClientTop       =   4440
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4230
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   2040
      TabIndex        =   9
      Top             =   3240
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid grdStandardDataFormat 
      Height          =   1335
      Left            =   240
      TabIndex        =   8
      Top             =   600
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2355
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin VB.TextBox txtDataFormat 
      Height          =   315
      Left            =   2280
      TabIndex        =   5
      Top             =   2400
      Width           =   3135
   End
   Begin VB.ComboBox cboDataTypes 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   5295
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "&Insert"
         Height          =   495
         Left            =   3600
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Predefined data formats"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Please enter a data format:"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Standard Data types:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
End
Attribute VB_Name = "frmStandardFormatDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1997-9. All Rights Reserved
'   File:       StandardFormatDefinition.bas
'   Author:     Mo Morris, July 1997
'   Purpose:    Used in library maintenance to maintain standard formats
'               for data items
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
 '  1       Mo Morris               4/07/97
'   2       Mo Morris               4/07/97
'   3       Andrew Newbigging       11/07/97
'   4       Mo Morris               17/07/97
'   5       Mo Morris               18/09/97
'   6       Mo Morris               8/10/98     gsMacro_DATABASE replaced by guser.DataBasePath
'                                               Changed look of form and added captions to
'                                               label to specify the datatype for the chosen format.
'   7      Will Casey               21/08/99    Added Datatypes to Combobox for user to choose datatype
'                                               from.Added module AddNewStandardFormatId to get a new Id
'                                               Id number and insert it in the database.Made the ID columns
'                                               invisible to the users.
'   8       WillC                   20/9/99     Changed the grid recordset to do a join across 2
'                                               tables due to a database change.
'   PN  20/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   9       WillC                   21/9/99     Changed If/else to a case statement
 '                                               removed multimedia and Category from combobox
'   10      WillC                   10/11/99    Added the error handlers
'   Mo Morris   16/11/99    DAO to ADO conversion
'   NCJ 8 Dec 99 - Bug fix to database insert ; tidying up field validation
'   WC 8/12/99  - Change to flexgrid
'   NCJ 9 Dec 99 - Give confirming message when deleting
'                   Match on Format and DataType when deleting
'   NCJ 14 Dec 99 - Check for empty data type in combo
'   TA 29/04/2000   removed subclassing
'   TA 08/06/2000   window border style chaged to fixxed single
'   NCJ 13/9/00 - Added LabTest to data formats combo
'   Mo 19/12/00 LabTest taken out of Standard formats.
'   Standarad formats will be set on ClinicalTests instead
'-----------------------------------------------------------------------------------'
'---------------------------------------------------------------------
'   FORM: frmStandardFormatDefinition
'
'   External references:
'   gnNewStandradFormatId
'   SetHelpContextId
'   ValidateDataFormat
'---------------------------------------------------------------------

Option Explicit


'SDM 28/01/00 SR2318    Not needed anymore
''---------------------------------------------------------------------
'Private Sub cboDataTypes_Change()
''---------------------------------------------------------------------
'' This can happen if they hit the Delete key to clear the selection
'' NCJ 14/12/99 - SR 2318
''---------------------------------------------------------------------
'
'    If cboDataTypes.Text > "" Then
'        ' Process like a click
'        Call cboDataTypes_Click
'    Else
'        cmdInsert.Enabled = False
'    End If
'
'End Sub

'---------------------------------------------------------------------
Private Sub cboDataTypes_Click()
'---------------------------------------------------------------------
' They've changed their data type selection
' so validate the current data format
'---------------------------------------------------------------------

    If Trim(txtDataFormat.Text) > "" Then
        If IsFormatValid Then
            cmdInsert.Enabled = True
        Else
            cmdInsert.Enabled = False
        End If
    Else
        cmdInsert.Enabled = False
    End If
    
End Sub
'---------------------------------------------------------------------
Private Sub cmdDelete_Click()
'---------------------------------------------------------------------
'Delete the relevant entry
'---------------------------------------------------------------------

Dim nRow As Integer
Dim sMsg As String
Dim sDataFormat As String
Dim sDataType As String
Dim nDataType As Integer

    On Error GoTo ErrHandler
          
    nRow = grdStandardDataFormat.Row
    grdStandardDataFormat.Col = 1   ' The format text column
    sDataFormat = grdStandardDataFormat.Text
    grdStandardDataFormat.Col = 2   ' The data type column
    sDataType = grdStandardDataFormat.Text
    sMsg = "Are you sure you want to delete the " & sDataType & " data format" _
                & vbCrLf & sDataFormat & " ?"
    
    Select Case MsgBox(sMsg, vbQuestion + vbYesNo, "Standard Data Formats")
        Case vbYes
            Call grdStandardDataFormat.RemoveItem(nRow)
            nDataType = GetSelectedDataType(sDataType)
            Call DeleteDataFormat(sDataFormat, nDataType)
        Case vbNo
            ' Do nothing
    End Select
    
Exit Sub
ErrHandler:
    Select Case Err.Number
       '  This error is ok as we wish to allow the user to delete the last userprofile
       '  in the grid and not encounter any problems.
        Case 30015
             Call DeleteDataFormat(sDataFormat, nDataType)
            grdStandardDataFormat.Clear
           Unload Me
        Case Else
              Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdDelete_Click")
                 Case OnErrorAction.Ignore
                     Resume Next
                 Case OnErrorAction.Retry
                     Resume
                 Case OnErrorAction.QuitMACRO
                     Call ExitMACRO
             End Select
    End Select
    

End Sub


Private Sub grdStandardDataFormat_Click()
    grdStandardDataFormat.Col = 0
    grdStandardDataFormat.ColSel = grdStandardDataFormat.Cols - 1
End Sub


'SDM 28/01/00 SR2318    Not needed anymore
''------------------------------------------------------------------------------------'
'Private Sub cboDataTypes_KeyPress(KeyAscii As Integer)
''------------------------------------------------------------------------------------'
'' Disallow the editing of combobox items
''------------------------------------------------------------------------------------'
'
'    If KeyAscii <> 0 Then
'        KeyAscii = 0
'    End If
'
'End Sub

'------------------------------------------------------------------------------------'
Private Sub Form_Load()
'------------------------------------------------------------------------------------'
' Load the Grid ,the combobox and disable the relevant Buttons and dont allow
' the editing of Datatypes through the grid
'------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler
    
    Me.Icon = frmMenu.Icon
    
    cmdInsert.Enabled = False
    
        With grdStandardDataFormat
        .Top = 200
        .Left = 200
        .Cols = 3
        .MergeCol(0) = True
        .Row = 0
        .ColWidth(0) = 200
        .Col = 0
        .Text = ""
        .ColWidth(1) = 2000
        .Col = 1
        .Text = "Data format"
        .ColWidth(2) = 2000
        .Col = 2
        .Text = "Data type name"
        .SelectionMode = flexSelectionByRow
        .ColAlignment(1) = flexAlignRightCenter
    End With

    RefreshGrid
    grdStandardDataFormat.Row = 1
    Call grdStandardDataFormat_Click
    
    cboDataTypes.AddItem ("Text")
    'cboDataTypes.AddItem ("Category")
    cboDataTypes.AddItem ("Integer Number")
    cboDataTypes.AddItem ("Real Number")
    cboDataTypes.AddItem ("Date/Time")
    'cboDataTypes.AddItem ("Multimedia")
    'Mo 19/12/00 LabTest taken out of Standard formats.
    'Standarad formats will be set on ClinicalTests instead
    'cboDataTypes.AddItem ("LabTest")    ' NCJ 13/9/00
    
    ' Initialise to "Text"
    cboDataTypes.ListIndex = 0
    

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

'------------------------------------------------------------------------------------'
Private Sub DeleteDataFormat(sDataFormat As String, nDataType As Integer)
'------------------------------------------------------------------------------------'
'delete the relevant entry
'------------------------------------------------------------------------------------'

Dim sSQL As String

    On Error GoTo ErrHandler
    
    sSQL = " DELETE  FROM StandardDataFormat " _
      & " WHERE DataFormat = '" & sDataFormat & "'" _
      & " AND DataTypeId = " & nDataType
     MacroADODBConnection.Execute sSQL, , adCmdText

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "DeleteDataFormat")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

    
End Sub


'----------------------------------------------------------------------------------------'
'Private Sub grdStandardDataFormat_MouseDown(Button As Integer, Shift As Integer, x As Single, _
'                                        y As Single)
''----------------------------------------------------------------------------------------'
'' if the selected column is the ValidationAction Col then disallow editing
'' else allow the edit
''----------------------------------------------------------------------------------------'
'Dim nColValue As Integer
'
'    On Error GoTo ErrHandler
'
'    ColValue = grdStandardDataFormat.ColContaining(x)
'
'    If nColValue = 1 Then
'        grdStandardDataFormat.AllowUpdate = False
'    Else
'        grdStandardDataFormat.AllowUpdate = True
'        If grdStandardDataFormat.Text = "" Then
'            MsgBox " Sorry, you may not leave a data format empty", vbOKOnly, "Validation Type"
'        End If
'    End If
'
'Exit Sub
'ErrHandler:
'    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
'                                    "grdStandardDataFormat_MouseDown")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            End
'    End Select
'
'End Sub

'----------------------------------------------------------------------------------------'
Private Sub txtDataFormat_Change()
'----------------------------------------------------------------------------------------'
' Check the length of the string to be less than 255 and not an empty string
'----------------------------------------------------------------------------------------'
Dim sText As String

    On Error GoTo ErrHandler
    
    sText = Trim(txtDataFormat.Text)
    
    If sText = "" Then
        txtDataFormat.Tag = sText
        cmdInsert.Enabled = False
        
    ElseIf Len(txtDataFormat.Text) > 255 Then
        MsgBox "A data format may not be more than 255 characters", vbOKOnly
        cmdInsert.Enabled = False
        txtDataFormat.Text = txtDataFormat.Tag
        
    ElseIf Not gblnValidString(sText, valOnlySingleQuotes) Then
        ' NCJ 31/5/00 - Include tildes in message
        MsgBox "Data formats may not contain quotes, tildes or pipe characters", vbOKOnly
        cmdInsert.Enabled = False
        txtDataFormat.Text = txtDataFormat.Tag
    
    Else
        txtDataFormat.Tag = sText
        cmdInsert.Enabled = True
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtDataFormat_Change")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Sub


'------------------------------------------------------------------------------------'
Private Sub cmdExit_Click()
'------------------------------------------------------------------------------------'

    Unload Me

End Sub

'------------------------------------------------------------------------------------'
Public Sub cmdInsert_Click()
'------------------------------------------------------------------------------------'
' First check to see if the format is valid if so add the new format
'------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler
    
    If IsFormatValid Then   ' Is format string valid?
    
        If AddNewStandardFormatId Then  ' Did it add successfully?
        
            'Set the text box to nothing
             txtDataFormat.Text = ""
             ' Reset the combo to "Text"
             cboDataTypes.ListIndex = 0
             
            'Show the insert in the grid immediately
            RefreshGrid
            
        End If
    
    End If
    
    ' Reset Insert button
    cmdInsert.Enabled = False
    
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

'------------------------------------------------------------------------------------'
Private Function IsFormatValid() As Boolean
'------------------------------------------------------------------------------------'
'Validate the dataformat
'------------------------------------------------------------------------------------'
Dim sFormat As String
Dim nDataItemType As Integer

    On Error GoTo ErrHandler

    ' NCJ 8 Dec 99 - Use GetSelectedDataType
    nDataItemType = GetSelectedDataType(cboDataTypes.Text)

    sFormat = Trim(txtDataFormat.Text)
    
    IsFormatValid = frmDataDefinition.ValidateDataFormat(sFormat, nDataItemType, True)

Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "IsFormatValid")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select

    
End Function

'------------------------------------------------------------------------------------'
Private Function GetSelectedDataType(sDataType As String) As Integer
'------------------------------------------------------------------------------------'
' Get the integer data type corresponding to the text data type
'------------------------------------------------------------------------------------'
    
    Select Case sDataType
    Case "Text"
        GetSelectedDataType = DataType.Text
    Case "Category"
        GetSelectedDataType = DataType.Category
    Case "Integer Number"
        GetSelectedDataType = DataType.IntegerData
    Case "Real Number"
        GetSelectedDataType = DataType.Real
    Case "Date/Time"
        GetSelectedDataType = DataType.Date
    Case "MultiMedia"
        GetSelectedDataType = DataType.Multimedia
    ' NCJ 13/9/00
    'Mo 19/12/00 LabTest taken out of Standard formats.
    'Case "LabTest"
    '    GetSelectedDataType = DataType.LabTest
    Case Else   ' Shouldn't happen!!
        GetSelectedDataType = DataType.Text
    End Select

End Function

'------------------------------------------------------------------------------------'
Private Function AddNewStandardFormatId() As Boolean
'------------------------------------------------------------------------------------'
' Get the Datatype and datatypeName depending on what the user has chosen
' from the cboDataTypes combobox
' Return TRUE if it was OK, or FALSE if already exists
'------------------------------------------------------------------------------------'
Dim rsTemp As ADODB.Recordset         'temporary recordset
Dim sDataFormat As String
Dim sSQL As String
Dim nDataTypeId As Integer
Dim nNewId As Integer

    On Error GoTo ErrHandler

    ' We'll set it to true when we've succeeded
    AddNewStandardFormatId = False
    
    sDataFormat = Trim(txtDataFormat.Text)
    
    ' Get the Datatype depending on what the user has chosen
    ' from the cboDataTypes combobox
    ' NCJ 8 Dec 99 - Use GetSelectedDataType
    nDataTypeId = GetSelectedDataType(cboDataTypes.Text)

    ' NCJ 9 Dec 99 - Check to see if data format already exists
    sSQL = " SELECT * FROM StandardDataFormat " _
            & " WHERE DataFormat = '" & sDataFormat & "'" _
            & " AND DataTypeId = " & nDataTypeId
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If rsTemp.EOF Then      ' It's not already there
    
        rsTemp.Close
        
        'Retrieve maximum StandardFormatId add 1 so you  have an incrementing Id each time
        'if there is no existing SDFId ie null then set it to 1
        
        sSQL = " SELECT (max(StandardDataFormatId) + 1) as NewStandardFormatId " _
                & "FROM StandardDataFormat "
            
        Set rsTemp = New ADODB.Recordset
        rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        ' NCJ - Set nNewId
        If IsNull(rsTemp!NewStandardFormatId) Then    'if no records exist
            nNewId = 1
        Else
            nNewId = rsTemp!NewStandardFormatId
        End If
     
        ' NCJ 8 Dec 99 - Bug fix to use correct value for DataTypeId SR
        sSQL = " INSERT INTO  StandardDataFormat" _
            & "(StandardDataFormatId,DataFormat,DataTypeId)" _
            & " VALUES(" & nNewId & ",'" & sDataFormat & "'," & nDataTypeId & ")"
                      
        MacroADODBConnection.Execute sSQL
        AddNewStandardFormatId = True
    Else
        ' It already exists
        MsgBox "This format already exists", vbOKOnly
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

Exit Function
ErrHandler:
 Select Case Err.Number
        Case 3022
             MsgBox "Sorry, you have tried to enter a format which already exists.", vbInformation, "Standard Data Formats"
        Case Else
            Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                            "AddNewStandardFormatId")
                Case OnErrorAction.Ignore
                    Resume Next
                Case OnErrorAction.Retry
                    Resume
                Case OnErrorAction.QuitMACRO
                    Call ExitMACRO
                    End
            End Select
    End Select

End Function

'------------------------------------------------------------------------------------'
Private Sub RefreshGrid()
'------------------------------------------------------------------------------------'
' Create the recordset doing the join between the StandardDataFormat and Datatype
' tables and load it into the grid
' Mo Morris   16/11/99    As part of the DAO to ADO conversion this Sub had to be
' re-written because it was not possible to pass an ADO recordset to a data control.
'------------------------------------------------------------------------------------'
Dim sSQL As String
Dim nRow As Long
Dim rsDatatypes As ADODB.Recordset

    On Error GoTo ErrHandler
    
'   ATN 21/12/99
'   Replaced INNER JOIN with WHERE clause
    sSQL = "SELECT DataFormat,DataTypeName" _
        & " FROM StandardDataFormat, DataType " _
        & " WHERE StandardDataFormat.DataTypeId = DataType.DataTypeId " _
        & " AND   StandardDataFormat.DataTypeId = DataType.DataTypeId "

    Set rsDatatypes = New ADODB.Recordset
    rsDatatypes.CursorLocation = adUseClient
    rsDatatypes.Open sSQL, MacroADODBConnection, adOpenKeyset, , adCmdText
    
    nRow = 1
    Do While Not rsDatatypes.EOF
    
        grdStandardDataFormat.Rows = nRow + 1
        grdStandardDataFormat.Row = nRow
        grdStandardDataFormat.Col = 0
        grdStandardDataFormat.Text = ""
        grdStandardDataFormat.CellBackColor = grdStandardDataFormat.BackColor
        grdStandardDataFormat.Col = 1
        grdStandardDataFormat.Text = rsDatatypes!DataFormat
        grdStandardDataFormat.Col = 2
        grdStandardDataFormat.Text = rsDatatypes!DataTypeName
    
        nRow = nRow + 1
        rsDatatypes.MoveNext
    Loop

    grdStandardDataFormat.Visible = True

    Set rsDatatypes = Nothing

    Exit Sub
    
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "RefreshGrid")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select

    
End Sub

'------------------------------------------------------------------------------------'
Private Function GetDataFormats() As ADODB.Recordset
'------------------------------------------------------------------------------------'
' Get the data for the grid
' Mo Morris   16/11/99    No longer called by Sub RefreshGrid
'------------------------------------------------------------------------------------'
Dim sSQL As String

    On Error GoTo ErrHandler
'   ATN 21/12/99
'   Oracle doesn't like INNER JOIN - replaced by WHERE clause
    sSQL = "SELECT DataFormat,DataTypeName" _
        & " FROM StandardDataFormat, DataType " _
        & " WHERE StandardDataFormat.DataTypeId = DataType.DataTypeId " _
        & " AND   StandardDataFormat.DataTypeId = DataType.DataTypeId "
    Set GetDataFormats = New ADODB.Recordset
    GetDataFormats.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "GetDataFormats")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Function

'------------------------------------------------------------------------------------'
Private Sub txtDataFormat_KeyPress(KeyAscii As Integer)
'------------------------------------------------------------------------------------'
' Accept RETURN as cmdInsert
'------------------------------------------------------------------------------------'

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If cmdInsert.Enabled Then
            Call cmdInsert_Click
        End If
    End If

End Sub
