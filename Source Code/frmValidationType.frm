VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmValidationType 
   Caption         =   "Validation Types   "
   ClientHeight    =   4245
   ClientLeft      =   11550
   ClientTop       =   3345
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   5910
   Begin VB.Frame fraType 
      Caption         =   "Existing types"
      Height          =   2415
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   5775
      Begin MSComctlLib.ListView lvwType 
         Height          =   2055
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   3625
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
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   4440
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   375
         Left            =   4440
         TabIndex        =   7
         Top             =   1920
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4500
      TabIndex        =   4
      Top             =   3780
      Width           =   1215
   End
   Begin VB.Frame fraUpdate 
      Caption         =   "Add/Change validation"
      Height          =   1215
      Left            =   60
      TabIndex        =   0
      Top             =   2460
      Width           =   5775
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Add"
         Height          =   375
         Left            =   4440
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   4440
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox cboAction 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtType 
         Height          =   285
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Validation name"
         Height          =   375
         Left            =   180
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Behaviour"
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmValidationType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999. All Rights Reserved
'   File:       frmValidationType.frm
'   Author:     Will Casey August 1999
'               Toby Aldridge JAn 2000
'   Purpose:    Maintains the table of Validation types by the Library manager
' Allows the maintenance of the ValidationType table. This form is available
' from the Parameters menu option on the frmMenu form. This form allows the
' insertion of data to the ValidationType table in the Macro database
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:

'   20/9/99     willC   Changed the RefreshGrid Proc to handle the join across two tables
'                       due to a database change.
'   PN  20/09/99        Added moFormIdleWatch object to handle system idle timer resets
'   21/9/99    willc    changed If/Else to Case statement
'   11/10/99   WillC    Added the error handlers
'   Mo Morris   15/11/99    DAO to ADO conversion
'   NCJ 7 Dec 99    Do not allow non-alpha chars in name
'   16/03/2000  TA  Redesigned form and behaviour, addedlistview and databse updating by ID
'   TA 7/6/2000 SR3572: have to dim VB6 listitem explicitly and listview default view changed to report
'   ASH 10/2/2003 Changed labelEdit property of lvwtype listview to lvwManual to avoid editting in listview
'------------------------------------------------------------------------------------'

Option Explicit
Option Base 0
Option Compare Binary

Const msSQL_TYPE = "SELECT ValidationType.ValidationTypeId, ValidationType.ValidationTypeName, ValidationType.ValidationActionId, ValidationAction.ValidationActionName FROM ValidationAction, ValidationType where ValidationAction.ValidationActionID = ValidationType.ValidationActionId"
Const msSQL_ACTION = "SELECT ValidationActionid, ValidationActionName FROM ValidationAction"
Const msSQL_MAX = "SELECT (MAX(ValidationTypeId) + 1) as NewValidationTypeId FROM ValidationType"

Const mnTYPE_ID = 1
Const mnTYPE_NAME = 2
Const mnTYPE_ACTION_ID = 3
Const mnTYPE_ACTION_NAME = 4


Const mnACTION_ID = 1
Const mnACTION_NAME = 2

Const mnMODE_SELECT = 1
Const mnMODE_ADD = 2
Const mnMODE_EDIT = 33


'edit/add/select mode
Private mlMode As Long
'initial form size
Private mlWidth As Long
Private mlHeight As Long

'----------------------------------------------------------------------------------------'
Private Sub cmdCancel_Click()
'------------------------------------------------------------------------------------'
' Cancel Clicked
'------------------------------------------------------------------------------------'
    
    On Error GoTo ErrHandler
    
    Call Mode(mnMODE_SELECT)
    lvwType_ItemClick lvwType.SelectedItem
Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdCancel_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Sub

'------------------------------------------------------------------------------------'
Private Sub cmdEdit_Click()
'------------------------------------------------------------------------------------'
' Changes the Validation Type Name
'------------------------------------------------------------------------------------'
   
   On Error GoTo ErrHandler
        
   Call Mode(mnMODE_EDIT)
   txtType.SetFocus
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
            Call MACROEnd
    End Select
    
End Sub

Private Sub cmdNew_Click()
'------------------------------------------------------------------------------------'
' Adds new Validation Name
'------------------------------------------------------------------------------------'

   On Error GoTo ErrHandler
   
   Call Mode(mnMODE_ADD)
   txtType.SetFocus
   txtType.Text = ""
   txtType.Tag = ""
   cboAction.ListIndex = 0
Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdNew_Click")
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
Private Sub Form_Resize()
'---------------------------------------------------------------------
' handle form resizing
'---------------------------------------------------------------------

    On Error GoTo ErrHandler:
    
    If Me.WindowState <> vbMinimized Then
        If Me.Width >= mlWidth Then
            'greater than initial form width
            fraType.Width = Me.ScaleWidth - 120
            fraUpdate.Width = fraType.Width
            lvwType.Width = fraType.Width - cmdEdit.Width - 360
            
            cmdExit.Left = fraType.Left + fraType.Width - cmdExit.Width - 120
            cmdEdit.Left = lvwType.Left + lvwType.Width + 120
            cmdNew.Left = cmdEdit.Left
            cmdUpdate.Left = cmdEdit.Left
            cmdCancel.Left = cmdEdit.Left
        Else
'            Me.Width = mlWidth
        End If
        
        If Me.Height >= mlHeight Then
            'greater than inital form height
            fraType.Height = Me.ScaleHeight - fraUpdate.Height - cmdExit.Height - 240
            lvwType.Height = fraType.Height - 360
            fraUpdate.Top = fraType.Top + fraType.Height + 60
           
            cmdExit.Top = Me.ScaleHeight - cmdExit.Height - 60
            
            cmdCancel.Top = fraUpdate.Height - cmdCancel.Height - 120
            cmdUpdate.Top = cmdCancel.Top - cmdUpdate.Height - 120
            cmdNew.Top = fraType.Height - cmdNew.Height - 120
            cmdEdit.Top = cmdNew.Top - cmdEdit.Height - 120
        Else
'            Me.Height = mlHeight
        End If
    End If
    
ErrHandler:

End Sub




'----------------------------------------------------------------------------------------'
Private Sub cmdExit_Click()
'----------------------------------------------------------------------------------------'
    
    Unload Me

End Sub

'----------------------------------------------------------------------------------------'
Private Sub cmdUpdate_Click()
'----------------------------------------------------------------------------------------'
' Get the max id in the the table add one to it then do the insert
'----------------------------------------------------------------------------------------'
Dim rsMax As ADODB.Recordset         'temporary recordset
'TA 7/6/2000 SR3572: have to dim VB6 control explicitly
Dim itmListItem As MSComctlLib.ListItem
Dim sType As String
Dim nTypeID As Integer
Dim nActionId As Integer

    On Error GoTo ErrHandler
    
    'get type and action
    sType = txtType.Text
    'check for duplicate name
    For Each itmListItem In lvwType.ListItems
        If UCase(itmListItem.Text) = UCase(sType) Then
            'duplicate found, is it the one we're editing
            If Right(itmListItem.Tag, InStrRev(itmListItem.Tag, ",") - 1) <> Right(lvwType.SelectedItem.Tag, InStrRev(lvwType.SelectedItem.Tag, ",") - 1) Then
                MsgBox "Duplicate validation name"
                Exit Sub
            End If
        End If
    Next
    nActionId = cboAction.ListIndex
    Select Case mlMode
    Case mnMODE_ADD
        'get new max value
        Set rsMax = New ADODB.Recordset
        rsMax.Open msSQL_MAX, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
        If IsNull(rsMax!NewValidationTypeId) Then    'if no records exist
            nTypeID = 1
        Else
            nTypeID = rsMax!NewValidationTypeId
        End If
        rsMax.Close
        Set rsMax = Nothing
        ' Place the new insert in the table immediately
        ' and initialise for next one
        Call InsertValidationType(nTypeID, sType, nActionId)
    Case mnMODE_EDIT
        'type id is number after comma in tag
        nTypeID = Right(lvwType.SelectedItem.Tag, InStrRev(lvwType.SelectedItem.Tag, ",") - 1)
        Call UpdateValidationType(nTypeID, sType, nActionId)
    End Select
    'refresh listview
    Call TypeRefresh
    Call Mode(mnMODE_SELECT)
    cboAction.ListIndex = 0
    txtType.Text = ""
    txtType.Tag = ""
    cmdEdit.Enabled = False
Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdUpdate_Click")
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
Private Sub UpdateValidationType(nValidationTypeId As Integer, sValidationTypeName As String, _
                                            nValidationActionId As Integer)
'----------------------------------------------------------------------------------------'
' Do the insert
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim ntemp As Integer
    
    On Error GoTo ErrHandler
    
    sSQL = "UPDATE ValidationType" _
            & " SET ValidationTypeName = '" & sValidationTypeName & "', ValidationActionID = " & nValidationActionId _
            & " WHERE ValidationTypeID = " & nValidationTypeId
    Set rsTemp = MacroADODBConnection.Execute(sSQL, ntemp, adAsyncFetch)
Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "UpdateValidationType")
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
Private Sub InsertValidationType(nValidationTypeId As Integer, sValidationTypeName As String, _
                                            nValidationActionId As Integer)
'----------------------------------------------------------------------------------------'
' Do the insert
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim ntemp As Integer
    
    On Error GoTo ErrHandler
    
    sSQL = " INSERT INTO ValidationType" _
            & "(ValidationTypeId,ValidationTypeName,ValidationActionId)" _
            & " VALUES (" & nValidationTypeId & ",'" & sValidationTypeName & "'," _
            & nValidationActionId & ")"
     Set rsTemp = MacroADODBConnection.Execute(sSQL, ntemp, adAsyncFetch)
Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "InsertValidationType")
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
Private Sub Form_Load()
'----------------------------------------------------------------------------------------'
' Get the data for the grid, load the listbox
'----------------------------------------------------------------------------------------'
Dim rsAction As ADODB.Recordset
    
    On Error GoTo ErrHandler
    
    Me.Icon = frmMenu.Icon
    
    'get inital form size
    mlWidth = Me.Width
    mlHeight = Me.Height
    FormCentre Me
    'fill combobox with valid behaviour actions
    Set rsAction = New ADODB.Recordset
    rsAction.CursorLocation = adUseClient
    rsAction.Open msSQL_ACTION, MacroADODBConnection, adOpenKeyset, , adCmdText
    Do While Not rsAction.EOF
        cboAction.AddItem rsAction.Fields(mnACTION_NAME - 1).Value, rsAction.Fields(mnACTION_ID - 1).Value
        rsAction.MoveNext
    Loop
    'fill listview with action name/behaviour
    Call TypeRefresh
    Call Mode(mnMODE_SELECT)
    cmdEdit.Enabled = False

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

Private Sub lvwType_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    On Error GoTo ErrHandler
    
    txtType.Text = Item.Text
    'action id is number before comma in tag
    cboAction.ListIndex = Left(Item.Tag, InStr(Item.Tag, ",") - 1)
    cmdEdit.Enabled = True
    Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "lvwType_ItemClick")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

Private Sub txtType_Change()
'----------------------------------------------------------------------------------------'
' If nothing has been entered in the text box or chosen from the listbox then
' leave the insert button disabled.
' NCJ 7 Dec 99 - Block non-alpha chars in name
'----------------------------------------------------------------------------------------'
Dim sTemp As String
Dim lSel As Long
    
    On Error GoTo ErrHandler
    
    lSel = txtType.SelStart
    sTemp = txtType.Text
    ' Check the text and only enable the Insert button
    ' if it's a valid name and there's an ActionType selected
    If sTemp = "" Then
        cmdUpdate.Enabled = False
    ElseIf Not IsValidValidationTypeName(sTemp) Then
        cmdUpdate.Enabled = False
        txtType.Text = txtType.Tag
        txtType.SelStart = lSel
    Else
        ' Store this valid value
        txtType.Tag = sTemp
        If mlMode <> mnMODE_SELECT Then
            cmdUpdate.Enabled = True
        End If
    End If
    Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtType_Change")
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
Private Function IsValidValidationTypeName(sName As String) As Boolean
'----------------------------------------------------------------------------------------'
' Return TRUE if text is valid ValidationTypeName
' Displays any necessary messages
'----------------------------------------------------------------------------------------'
    
    On Error GoTo ErrHandler
    
    IsValidValidationTypeName = False
    If sName > "" Then
        If Not gblnValidString(sName, valOnlySingleQuotes) Then
            MsgBox "Name" & gsCANNOT_CONTAIN_INVALID_CHARS, _
                    vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
        ElseIf Not gblnValidString(sName, valAlpha + valSpace) Then
            MsgBox "Name may only contain letters", _
                    vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
        Else
            IsValidValidationTypeName = True
        End If
    End If
Exit Function

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "IsValidValidationTypeName")
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
Private Sub Mode(lMode As Long)
'------------------------------------------------------------------------------------'
' Enable/disable controls according to situation
'------------------------------------------------------------------------------------'
    
    On Error GoTo ErrHandler
    
    Select Case lMode
    Case mnMODE_SELECT
        lvwType.Enabled = True
        cmdEdit.Enabled = True
        cmdNew.Enabled = True
        txtType.Enabled = False
        cboAction.Enabled = False
        cmdUpdate.Enabled = False
        cmdCancel.Enabled = False
        fraUpdate.Caption = "Add/Change validation"
    Case mnMODE_EDIT
        lvwType.Enabled = False
        cmdEdit.Enabled = False
        cmdNew.Enabled = False
        txtType.Enabled = True
        cboAction.Enabled = True
        cmdUpdate.Enabled = True
        cmdUpdate.Caption = "&Change"
        cmdCancel.Enabled = True
        fraUpdate.Caption = "Change validation"
    Case mnMODE_ADD
        lvwType.Enabled = False
        cmdEdit.Enabled = False
        cmdNew.Enabled = False
        txtType.Enabled = True
        cboAction.Enabled = True
        cmdUpdate.Enabled = False
        cmdUpdate.Caption = "&Add"
        cmdCancel.Enabled = True
        fraUpdate.Caption = "Add validation"
    End Select
    mlMode = lMode
Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "Mode")
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
Private Sub TypeRefresh()
'------------------------------------------------------------------------------------'
' Refresh the type listbox
'------------------------------------------------------------------------------------'
Dim rsType As ADODB.Recordset
'TA 7/6/2000 SR3572: have to dim VB6 listitem explicitly
Dim itmNew As MSComctlLib.ListItem
Dim lTypeWidth As Long
Dim lActionWidth As Long

    On Error GoTo ErrHandler
    
    'clear listview
    lvwType.ColumnHeaders.Clear
    lvwType.ListItems.Clear
    'add column headings
    With lvwType.ColumnHeaders.Add
        .Text = "Validation name"
    End With
    With lvwType.ColumnHeaders.Add
        .Text = "Behaviour"
    End With
'    'get inital column widths  (adding blank pixels)
'    lTypeWidth = lvwType.Parent.TextWidth("Validation name") + 12 * Screen.TwipsPerPixelX
'    lActionWidth = lvwType.Parent.TextWidth("Behaviour") + 12 * Screen.TwipsPerPixelX
    'retrieve data
    Set rsType = New ADODB.Recordset
    rsType.Open msSQL_TYPE, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsType.EOF
        'check if columns should be widened
        If lvwType.Parent.TextWidth(rsType.Fields(mnTYPE_NAME - 1).Value) > lTypeWidth Then
            lTypeWidth = lvwType.Parent.TextWidth(rsType.Fields(mnTYPE_NAME - 1).Value)
        End If
        If lvwType.Parent.TextWidth(rsType.Fields(mnTYPE_ACTION_NAME - 1).Value) > lActionWidth Then
            lActionWidth = lvwType.Parent.TextWidth(rsType.Fields(mnTYPE_ACTION_NAME - 1).Value)
        End If
        'append next row to listview
        With lvwType.ListItems.Add(, , rsType.Fields(mnTYPE_NAME - 1).Value)
            .SubItems(1) = rsType.Fields(mnTYPE_ACTION_NAME - 1).Value
            'set the tag as type id and action id separated by a comma
            .Tag = rsType.Fields(mnTYPE_ACTION_ID - 1).Value & "," & rsType.Fields(mnTYPE_ID - 1).Value
        End With
        rsType.MoveNext
    Loop
    'set appropriate col widths
    lvwType.ColumnHeaders(1).Width = lTypeWidth + 6 * Screen.TwipsPerPixelX
    lvwType.ColumnHeaders(2).Width = lActionWidth + 6 * Screen.TwipsPerPixelX
    Set rsType = Nothing
    Exit Sub
    
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "TypeRefresh")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub
