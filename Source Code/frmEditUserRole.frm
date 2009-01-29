VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditUserRole 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MACRO Edit Existing Role"
   ClientHeight    =   3360
   ClientLeft      =   4845
   ClientTop       =   4500
   ClientWidth     =   7905
   Icon            =   "frmEditUserRole.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "&Select All"
      Height          =   495
      Left            =   3660
      TabIndex        =   10
      Top             =   2580
      Width           =   1215
   End
   Begin VB.CommandButton cmdDeselectAll 
      Caption         =   "&Deselect All"
      Height          =   495
      Left            =   6420
      TabIndex        =   9
      Top             =   2580
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   3255
      Left            =   3480
      TabIndex        =   7
      Top             =   0
      Width           =   4335
      Begin MSComctlLib.ListView lvwRoleFunctions 
         Height          =   2235
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   3942
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
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
   Begin MSComctlLib.ImageList Img1 
      Left            =   2160
      Top             =   -240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditUserRole.frx":0442
            Key             =   "BoxNoTick"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditUserRole.frx":089E
            Key             =   "BoxTick"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   3135
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Close"
         Height          =   495
         Left            =   1680
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "&Apply"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox txtRoleCode 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   3135
   End
   Begin VB.TextBox txtRoleDescription 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Role description:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Role code:"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmEditUserRole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       frmEditUserRole.frm
'   Author:     Will Casey
'   Purpose:
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
'   Revisions:
'   PN  20/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   WillC 11/10/99  Added the gLog call ,to log editing of the role being edited in
'                   the cmdApply sub.
'   WillC   13/12/99 added the ability to select and deselect all functions
'   TA 08/05/2000   removed subclassing
'--------------------------------------------------------------------------------

'---------------------------------------------------------------------

Option Explicit
Option Compare Binary
Option Base 0

Private suCodeCol As Collection

'------------------------------------------------------------------------------'
Private Sub cmdDeselectAll_Click()
'------------------------------------------------------------------------------'
'Uncheck all the icons in the listview
'------------------------------------------------------------------------------'

 Dim n As Integer
 
    For n = 1 To lvwRoleFunctions.ListItems.Count
        lvwRoleFunctions.ListItems.Item(n).SmallIcon = "BoxNoTick"
    Next n

End Sub

'------------------------------------------------------------------------------'
Private Sub cmdSelectAll_Click()
'------------------------------------------------------------------------------'
'Check all the icons in the listview
'------------------------------------------------------------------------------'
 Dim n As Integer
 
    For n = 1 To lvwRoleFunctions.ListItems.Count
        lvwRoleFunctions.ListItems.Item(n).SmallIcon = "BoxTick"
    Next n
    
End Sub

'------------------------------------------------------------------------------------'
Private Sub cmdApply_Click()
'------------------------------------------------------------------------------------'
'
'Find which databases have a tick against them and those that do insert them
'otherwise delete them , then get the functions against a role and do the same thing
' so you have a valid insert as configured by the Users choices.
' Added the gLog call to log editing of the role being edited
'------------------------------------------------------------------------------------'
Dim tmpFunctionCode As String
Dim tmpRoleCode As String
Dim n As Integer
Dim sRoleDescription As String

    On Error GoTo ErrHandler

    'sDatabaseDescription = frmUserMaintenance.lvwUserDatabase.SelectedItem
    'sUserCode = txtUserCode.Text
    tmpRoleCode = txtRoleCode.Text
    sRoleDescription = txtRoleDescription.Text
    If tmpRoleCode = "" Then
        MsgBox "Please enter a Role Description", vbOKOnly
    End If
    
    'Loop around the members in the listview where the icon is Boxtick add the
    'function to the database for the role, where the icon is BoxNoTick
    'remove the function from the role.
    
    For n = 1 To lvwRoleFunctions.ListItems.Count
        ' Get the FunctionCode using the Function Name for the database insert/delete
        ' using the collection set up for that.
        tmpFunctionCode = suCodeCol.Item(lvwRoleFunctions.ListItems.Item(n))
         ' Do the Inserts/Deletes now that we have all the fields necessary for the
         ' RoleFunction table
         ' if the icon is BoxNoTick remove the Function from a Role the RoleFuction table
         ' if the icon is BoxTick add the function to a Role in the RoleFuction table
        If lvwRoleFunctions.ListItems.Item(n).SmallIcon = "BoxNoTick" Then
            RemoveFunctionFromRole tmpRoleCode, tmpFunctionCode
        Else
            Call gblnHasRoleFunction(tmpRoleCode, tmpFunctionCode)
            ' Call AddFunctionToRole(tmpRoleCode, tmpFunctionCode)
        End If
    Next
    
    Call EditRole(tmpRoleCode, sRoleDescription)
    ' Update the Role table to show the changes
    RefreshRoles
    'Clear the Function listbox for neatness
    frmRoleMgmt.lstUserFunction.Clear
    ' log the change
    Call gLog("sCreateNewRole", "User role" & " " & tmpRoleCode & " " & "edited.")
    frmRoleMgmt.cmdEditRole.Enabled = False
        
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdApply_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
 End Sub

Private Sub cmdCancel_Click()

Unload Me

End Sub

'------------------------------------------------------------------------------------'
Private Sub Form_Load()
'------------------------------------------------------------------------------------'
'load the list of functions into the listview disable the apply button
'
'------------------------------------------------------------------------------------'
Dim clmX As MSComctlLib.ColumnHeader
Dim sRoleCode As String
Dim sRoleDescription As String
Dim rsRoleDescription As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler

    Me.KeyPreview = True
    txtRoleCode.Enabled = False
    cmdApply.Enabled = False
    
    sRoleCode = frmRoleMgmt.lstUserRole.Text
    txtRoleCode.Text = sRoleCode

    sSQL = "SELECT RoleDescription FROM Role WHERE RoleCode = '" & sRoleCode & "'"
    Set rsRoleDescription = New ADODB.Recordset
    rsRoleDescription.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, , adCmdText
    sRoleDescription = rsRoleDescription!RoleDescription
    Set rsRoleDescription = Nothing
    
    txtRoleDescription.Text = sRoleDescription
    lvwRoleFunctions.View = lvwReport ' Set View property to Report
    lvwRoleFunctions.SmallIcons = Img1 'The images for the listbox are imagelist Img1
    RefreshFunctions

    ' Add ColumnHeaders with appropriate widths to lvwRoleFunctions
    Set clmX = lvwRoleFunctions.ColumnHeaders.Add(, , "RoleFunctions", lvwRoleFunctions.Width - 650)

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "Form_Load")
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
Private Sub RefreshDatabases()
'------------------------------------------------------------------------------------'
' Add all databases to the combobox
'------------------------------------------------------------------------------------'
'Dim rsDatabases As ADODB.Recordset
'
'' Get the list of Databases
'    Set rsDatabases = gdsDatabaseList
'
'    ' While the record is not the last record, add a ListItem object.
'    With rsDatabases
'        Do Until .EOF = True
'          cboDatabases.AddItem .Fields(0).Value
'            'this places the databasenames in the listview
'                 .MoveNext   ' Move to next record.
'        Loop
'    End With
    
 End Sub


'------------------------------------------------------------------------------------'
Private Sub RefreshFunctions()
'------------------------------------------------------------------------------------'
'Load all the functions into the Listview and then tick the box for the ones that
' are present in the main form listbox.
'
'------------------------------------------------------------------------------------'
Dim itmX As MSComctlLib.ListItem
Dim rsFunctions As ADODB.Recordset
Dim n As Integer
Dim blnItMatches As Boolean

    On Error GoTo ErrHandler

    Set suCodeCol = New Collection
    
    ' Get the list of Functions
    Set rsFunctions = gdsFunctionList
    
    ' While the record is not the last record, load all
    ' ListItem objects into the listview.
    With rsFunctions
        Do Until .EOF = True  'Add the FunctionNames
          Set itmX = lvwRoleFunctions.ListItems.Add(, , .Fields(1).Value)
            itmX.SmallIcon = "BoxNoTick"
            ' Use the collection below to hold the functioncode for the relevant
            ' function name, for the inserts/deletes in the Apply click event.
            ' fields(0) is the function code, fields(1) is the function name
            suCodeCol.Add .Fields(0).Value, .Fields(1).Value
            ' Loop around the Function names present in the frmUserMaintenance function listbox
            ' and then tick the corresponding functions in the function listview on this form.
            For n = 0 To frmRoleMgmt.lstUserFunction.ListCount
                ' Set them all to not present
                blnItMatches = False
                ' If the present item IS in the list then tick the function
                If itmX = frmRoleMgmt.lstUserFunction.List(n) Then
                    blnItMatches = True
                    If blnItMatches = True Then
                        itmX.SmallIcon = "BoxTick"
                    End If
                End If
            Next
        .MoveNext   ' Move to next record.
        Loop
    End With
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "RefreshFunctions")
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
Private Sub lvwRoleFunctions_Click()
'------------------------------------------------------------------------------------'
'
'This allows the toggling between ticked and not ticked, "default" is not ticked
'------------------------------------------------------------------------------------'
On Error GoTo ErrHandler
    
    If lvwRoleFunctions.SelectedItem.SmallIcon = "BoxNoTick" Then
        lvwRoleFunctions.SelectedItem.SmallIcon = "BoxTick"
    Else
        lvwRoleFunctions.SelectedItem.SmallIcon = "BoxNoTick"
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "lvwRoleFunctions_Click")
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
Private Sub RefreshRoles()
'------------------------------------------------------------------------------------'
'Add all the roles to the UserRole listbox..
'------------------------------------------------------------------------------------'
Dim rsRoles As ADODB.Recordset

    On Error GoTo ErrHandler
    
    frmRoleMgmt.lstUserRole.Clear
    ' Get the list of roles
    Set rsRoles = gdsRoleList
    With rsRoles
       Do Until .EOF = True
         'this adds the RoleCode field from the table to the list box
         frmRoleMgmt.lstUserRole.AddItem .Fields(0).Value
         .MoveNext
       Loop
    End With
    
    Unload Me
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "RefreshRoles")
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
Private Sub txtRoleCode_Change()
'------------------------------------------------------------------------------------'
' if the textboxes are empty leave the apply button disabled
'------------------------------------------------------------------------------------'
    On Error GoTo ErrHandler

    If txtRoleDescription.Text = "" Then
       cmdApply.Enabled = False
    Else
        cmdApply.Enabled = True
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtRoleCode_Change")
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
Private Sub txtRoleDescription_Change()
'------------------------------------------------------------------------------------'
' if the textboxes are empty leave the apply button disabled
' NCJ 15/12/99 - Only enable Apply if description is valid
'------------------------------------------------------------------------------------'
Dim sDescription As String

    On Error GoTo ErrHandler

    sDescription = Trim(txtRoleDescription.Text)
    If sDescription = "" Then
        cmdApply.Enabled = False
    Else
        cmdApply.Enabled = IsValidString(sDescription)
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtRoleDescription_Change")
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
Private Sub EditRole(sRoleCode As String, sRoleDescription As String)
'------------------------------------------------------------------------------------'
' If a roledescription is changed this handles it.
'------------------------------------------------------------------------------------'
Dim sSQL As String

    On Error GoTo ErrHandler
    
    sSQL = " UPDATE Role SET "
    '  & " RoleCode = " & "'" & sRoleCode & "',"
    sSQL = sSQL & " RoleDescription = " & "'" & sRoleDescription & "'"
    sSQL = sSQL & " WHERE RoleCode = '" & sRoleCode & "'"
    SecurityADODBConnection.Execute sSQL, , adCmdText
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "EditRole")
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
Private Function IsValidString(sDescription As String) As Boolean
'----------------------------------------------------------------------------------------'
' Return TRUE if text is valid name for Role description
' Displays any necessary messages
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler
    
    IsValidString = False
       
    If sDescription > "" Then
        If Not gblnValidString(sDescription, valOnlySingleQuotes) Then
            MsgBox " Role descriptions may not contain double or backward quotes or the | character", _
                    vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
        ElseIf Not gblnValidString(sDescription, valAlpha + valNumeric + valSpace) Then
            MsgBox " Role descriptions may only contain alphanumeric characters", _
                    vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
        ElseIf Len(sDescription) > 255 Then
            MsgBox " Role descriptions may not be more than 255 characters", _
                    vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
        Else
            IsValidString = True
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

