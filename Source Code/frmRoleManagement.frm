VERSION 5.00
Begin VB.Form frmRoleManagement 
   BorderStyle     =   0  'None
   Caption         =   "Role Management"
   ClientHeight    =   8205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraRole 
      Caption         =   "Role"
      Height          =   1385
      Left            =   60
      TabIndex        =   14
      Top             =   120
      Width           =   7755
      Begin VB.CheckBox chkSysAdmin 
         Alignment       =   1  'Right Justify
         Caption         =   "System &Administrator"
         Height          =   345
         Left            =   180
         TabIndex        =   17
         Top             =   1000
         Width           =   1875
      End
      Begin VB.TextBox txtRoleCode 
         Height          =   345
         Left            =   1860
         MaxLength       =   15
         TabIndex        =   1
         Top             =   180
         Width           =   1935
      End
      Begin VB.TextBox txtRoleDescription 
         Height          =   345
         Left            =   1860
         TabIndex        =   2
         Top             =   600
         Width           =   3915
      End
      Begin VB.Label lblRoleCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Role &Code"
         Height          =   255
         Left            =   330
         TabIndex        =   16
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label lblRoleDescription 
         Alignment       =   1  'Right Justify
         Caption         =   "Role &Description"
         Height          =   255
         Left            =   330
         TabIndex        =   15
         Top             =   660
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdRemoveAll 
      Caption         =   "<<"
      Height          =   315
      Left            =   3480
      TabIndex        =   8
      Top             =   6900
      Width           =   915
   End
   Begin VB.CommandButton cmdMoveAll 
      Caption         =   ">>"
      Height          =   315
      Left            =   3480
      TabIndex        =   7
      Top             =   6480
      Width           =   915
   End
   Begin VB.CommandButton cmdRemoveOne 
      Caption         =   "<"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3480
      TabIndex        =   6
      Top             =   4980
      Width           =   915
   End
   Begin VB.CommandButton cmdMoveOne 
      Caption         =   ">"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3480
      TabIndex        =   5
      Top             =   4560
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   6600
      TabIndex        =   11
      Top             =   7740
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   7740
      Width           =   1215
   End
   Begin VB.Frame fraSelectedRoles 
      Caption         =   "Selected Role Functions"
      Height          =   3600
      Left            =   4560
      TabIndex        =   13
      Top             =   4080
      Width           =   3285
      Begin VB.ListBox lstSelectedRoles 
         Height          =   3180
         Left            =   120
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   3045
      End
   End
   Begin VB.Frame fraAvailableRoles 
      Caption         =   "Available Role Functions"
      Height          =   3600
      Left            =   60
      TabIndex        =   12
      Top             =   4080
      Width           =   3285
      Begin VB.ListBox lstAvailableRoles 
         Height          =   3180
         Left            =   120
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   3045
      End
   End
   Begin VB.Frame fraMacroModules 
      Caption         =   "Select MACRO Module(s) to display the related Role functions"
      Height          =   2475
      Left            =   60
      TabIndex        =   0
      Top             =   1560
      Width           =   7755
      Begin VB.ListBox lstMACROModules 
         Height          =   2085
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   240
         Width           =   7515
      End
   End
End
Attribute VB_Name = "frmRoleManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2002-2008. All Rights Reserved
'   File:       frmRoleManagement.frm
'   Author:     Ashitei Trebi-Ollennu, September 2002
'   Purpose:    Adds/edits roles
'------------------------------------------------------------------------------
'REVISIONS:
'ASH 17/1/2003 - Fix for creating similar rolecodes in different character cases in sub UpdateRoles
'ASH 27/1/2003 - Added Batch Data Entry
' NCJ 4 Mar 03 - Added Batch Validation
' NCJ 18 Nov 04 - Issue 2424 - Disable OK button as soon as it's pressed
' NCJ 28 Feb 08 - Issue 3003 - Disable update buttons if user does not have Maintain Role permission
'-------------------------------------------------------------------------------

Option Explicit

Private mcolRoles As New Collection
Private mcolSystemManagement As Collection
Private mcolCreateDataViews As Collection
Private mcolStudyDefinition As Collection
Private mcolLibraryManagement As Collection
Private mcolDataEntry As Collection
Private mcolDataReview As Collection
Private mcolQueryModule As Collection
Private mcolBatchDataEntryModule As Collection
Private mcolBatchValidationModule As Collection

Private mbFormLoading As Boolean
Private msRoleCode As String
Private mbChanged As Boolean
Private mbNewRole As Boolean

' NCJ 28 Feb 08 - Issue 3003 - Store user's Maintain Role permission
Private mbCanUpdate As Boolean

Private Const msMACRO_SYSTEM_MNGT = "System Management"
Private Const msMACRO_DATA_ENTRY = "Data Entry"
Private Const msMACRO_CREATE_DATA_VIEWS = "Create Data Views"
Private Const msMACRO_QUERY_MODULE = "Query Module"
Private Const msMACRO_LIBRARY_MNGT = "Library Management"
Private Const msMACRO_STUDY_DEF = "Study Definition"
Private Const msMACRO_DATA_REVIEW = "Data Review"
Private Const msMACRO_BATCH_DATA_ENTRY = "Batch Data Entry"
Private Const msMACRO_BATCH_VALIDATION = "Batch Validation"

Private Const msMACRO_SYSTEM_MNGT_PREFIX = "SM"
Private Const msMACRO_DATA_ENTRY_PREFIX = "DE"
Private Const msMACRO_CREATE_DATA_VIEWS_PREFIX = "DV"
Private Const msMACRO_QUERY_MODULE_PREFIX = "QM"
Private Const msMACRO_LIBRARY_MNGT_PREFIX = "LM"
Private Const msMACRO_STUDY_DEF_PREFIX = "SD"
Private Const msMACRO_DATA_REVIEW_PREFIX = "DR"
Private Const msMACRO_BATCH_DATA_ENTRY_PREFIX = "BD"
Private Const msMACRO_BATCH_VALIDATION_PREFIX = "BV"

'--------------------------------------------------------------------------------------
Public Sub Display(Optional ByRef sRoleCode As String = "")
'--------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------
    
    mbFormLoading = True
    Load Me
    'MLM 03/12/02:
    Me.Top = frmMenu.Height 'to make sure that the form isn't visible until it has initialised
    'frmMenu.MDIResize
    'DoEvents
    
    ' NCJ 28 Feb 08 - Can user do any updates?
    mbCanUpdate = goUser.CheckPermission(gsFnMaintRole)
    
    cmdMoveOne.Enabled = False
    cmdRemoveOne.Enabled = False
    
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
    mbChanged = False
        
    'only enable check box if user is a sys admin
    chkSysAdmin.Enabled = goUser.SysAdmin
    
    ' Only enable Description if user can update
    txtRoleDescription.Locked = Not mbCanUpdate
    
    mbNewRole = (sRoleCode = "")
    msRoleCode = sRoleCode
    
    Call GetRoles
    
    Call LoadMACROModules

    lstMACROModules.Enabled = True
    
    EnableButtons
    'EnableOKButton
    
    
    Me.Icon = frmMenu.Icon
    
    'Me.Show
    
    mbFormLoading = False


End Sub

'------------------------------------------------------------------------------
Private Sub chkSysAdmin_Click()
'------------------------------------------------------------------------------

    mbChanged = True
    
    EnableOKButton
    
End Sub

'------------------------------------------------------------------------------
Private Sub cmdCancel_Click()
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------

    If mbChanged Then
        If DialogQuestion("Are you sure you want to lose the changes?", gsDIALOG_TITLE) = vbYes Then
            
            Call GetRoles

            LoadMACROModules
        
            lstMACROModules.Enabled = True
            
            EnableButtons
            'EnableOKButton
            mbChanged = False
            cmdOK.Enabled = False
            cmdCancel.Enabled = False
            cmdMoveOne.Enabled = False
            cmdRemoveOne.Enabled = False

        End If
    End If
    
End Sub

'----------------------------------------------------------------------------
Private Sub cmdMoveAll_Click()
'----------------------------------------------------------------------------
'calls routine to add all roles to the Selected roles listbox
'also enables and disables controls
'----------------------------------------------------------------------------
    
    mbChanged = True
    
    AddToSelectedRoles (1)
    cmdMoveAll.Enabled = False
    cmdMoveOne.Enabled = False
    EnableButtons
    EnableOKButton

End Sub

'------------------------------------------------------------------------------
Private Sub cmdMoveOne_Click()
'------------------------------------------------------------------------------
'calls routine to add a selected role(s) to the Selected roles listbox
'also enables
'------------------------------------------------------------------------------
    
    mbChanged = True
    
    AddToSelectedRoles (0)
    If lstAvailableRoles.ListCount = 0 Then
        cmdMoveOne.Enabled = False
        cmdMoveAll.Enabled = False
    End If
    EnableButtons
    EnableOKButton
    
End Sub

'-----------------------------------------------------------------------------
Private Sub cmdOK_Click()
'-----------------------------------------------------------------------------
' Apply the new role definition
'-----------------------------------------------------------------------------

    On Error GoTo ErrLabel

    ' NCJ 18 Nov 04 - Issue 2424 - Moved disablement to here (to prevent further clicking)
    cmdOK.Enabled = False
    txtRoleCode.Enabled = False
    cmdCancel.Enabled = False
    
    If mbChanged Then
        UpdateRoles
    End If
    
    mbNewRole = False
    mbChanged = False

Exit Sub
ErrLabel:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdOK_Click|frmRoleManagement")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select

End Sub

'------------------------------------------------------------------------------
Private Sub cmdRemoveAll_Click()
'------------------------------------------------------------------------------
'calls routine to remove all roles from Selected roles listbox to
'Available roles listbox
'also enables and disables controls
'------------------------------------------------------------------------------
    
    mbChanged = True
    
    RemoveFromSelectedRoles (1)
    cmdMoveAll.Enabled = True
    EnableButtons
    EnableOKButton
    
End Sub

'------------------------------------------------------------------------------
Private Sub cmdRemoveOne_Click()
'------------------------------------------------------------------------------
'calls routine to remove selected role(s) from lstSelected listbox
'also enables and disables controls
'------------------------------------------------------------------------------
    
    mbChanged = True

    RemoveFromSelectedRoles (0)
    cmdMoveAll.Enabled = True
    EnableButtons
    EnableOKButton
    
End Sub

'--------------------------------------------------------------------------------
Private Sub LoadMACROModules()
'--------------------------------------------------------------------------------
'Loads all the current MACRO Modules into a list box
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim rsModules As ADODB.Recordset
Dim sModule As String

    On Error GoTo ErrHandler
    
    lstMACROModules.Clear
    lstAvailableRoles.Clear

'    'get MACRO Modules from the security database
'    Set rsModules = New ADODB.Recordset
'    sSQL = "SELECT MACROFunction FROM MACROFunction"  ' WHERE Left(FunctionCode,2)='" & "F1" & "'"
'    rsModules.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
'
'    rsModules.MoveFirst
'    Do Until rsModules.EOF
'        If Left(rsModules!MACROFunction, 6) = "Access" Then
'            'if it is an access certain module function then trim off the prefix ACCESS and spaces left
'            sModule = Trim(Mid(rsModules!MACROFunction, 7))
'            'load roles for macro module
'            GetAllRoleFunctions (sModule)
'            lstMACROModules.AddItem sModule
'        End If
'        rsModules.MoveNext
'    Loop
    
    ' NCJ 4 Mar 03 - Add the modules explicitly, using the defined constants
    sModule = msMACRO_SYSTEM_MNGT
    Call GetAllRoleFunctions(sModule)
    lstMACROModules.AddItem sModule
    
    sModule = msMACRO_STUDY_DEF
    Call GetAllRoleFunctions(sModule)
    lstMACROModules.AddItem sModule
    
    sModule = msMACRO_LIBRARY_MNGT
    Call GetAllRoleFunctions(sModule)
    lstMACROModules.AddItem sModule
    
    sModule = msMACRO_DATA_ENTRY
    Call GetAllRoleFunctions(sModule)
    lstMACROModules.AddItem sModule
    
    sModule = msMACRO_DATA_REVIEW
    Call GetAllRoleFunctions(sModule)
    lstMACROModules.AddItem sModule
    
    sModule = msMACRO_CREATE_DATA_VIEWS
    Call GetAllRoleFunctions(sModule)
    lstMACROModules.AddItem sModule
    
    sModule = msMACRO_QUERY_MODULE
    Call GetAllRoleFunctions(sModule)
    lstMACROModules.AddItem sModule
    
    sModule = msMACRO_BATCH_DATA_ENTRY
    Call GetAllRoleFunctions(sModule)
    lstMACROModules.AddItem sModule
    
    sModule = msMACRO_BATCH_VALIDATION
    Call GetAllRoleFunctions(sModule)
    lstMACROModules.AddItem sModule
    
    ' Checks all modules. This is the default.
    CheckAllModules

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmRoleManagement.LoadMACROModules"
    
End Sub

'------------------------------------------------------------------------------
Private Sub Form_Resize()
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
    
    On Error Resume Next
            
    'role frame, and its contents
    fraRole.Width = Me.Width - 120
    
    'modules frame, and its contents
    fraMacroModules.Width = fraRole.Width
    lstMACROModules.Width = fraMacroModules.Width - 240
    
    'the buttons in between the function lists
    cmdMoveOne.Left = (Me.Width - cmdMoveOne.Width) \ 2
    cmdMoveAll.Left = cmdMoveOne.Left
    cmdRemoveOne.Left = cmdMoveOne.Left
    cmdRemoveAll.Left = cmdMoveOne.Left
    
    cmdMoveAll.Top = (Me.Height - cmdMoveAll.Top \ 4)
    cmdRemoveAll.Top = cmdMoveAll.Top + 400
    
    'the available roles frame, and its contents
    fraAvailableRoles.Width = cmdMoveOne.Left - 120
    fraAvailableRoles.Height = Me.Height - fraAvailableRoles.Top - cmdOK.Height - 240
    lstAvailableRoles.Width = fraAvailableRoles.Width - 240
    lstAvailableRoles.Height = fraAvailableRoles.Height - 360
    
    'the selected roles frame and its contents
    fraSelectedRoles.Left = cmdMoveOne.Left + cmdMoveOne.Width + 60
    fraSelectedRoles.Width = fraAvailableRoles.Width
    fraSelectedRoles.Height = fraAvailableRoles.Height
    lstSelectedRoles.Width = lstAvailableRoles.Width
    lstSelectedRoles.Height = lstAvailableRoles.Height
    
    'the OK and Cancel buttons
    cmdCancel.Top = Me.Height - cmdCancel.Height - 100
    cmdCancel.Left = Me.Width - cmdCancel.Width - 60
    cmdOK.Top = cmdCancel.Top
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 120
        
End Sub

'--------------------------------------------------------------------------------
Private Sub lstAvailableRoles_Click()
'--------------------------------------------------------------------------------
' NCJ 28 Feb 08 - Check user's update permission
'--------------------------------------------------------------------------------
Dim i As Integer
Dim nCount As Integer

    If lstAvailableRoles.ListIndex > -1 Then
        cmdMoveOne.Enabled = mbCanUpdate
    Else
        cmdMoveOne.Enabled = False
    End If
  
End Sub

'---------------------------------------------------------------------------------
Private Sub lstAvailableRoles_DblClick()
'---------------------------------------------------------------------------------
'adds role(s) to listbox
'---------------------------------------------------------------------------------

    ' NCJ 28 Feb 08 - Check user's role update permission
    If Not mbCanUpdate Then Exit Sub
    
    AddToSelectedRoles (0)
    EnableButtons
    EnableOKButton
    mbChanged = True

End Sub

'-----------------------------------------------------------------------------------
Private Sub lstAvailableRoles_GotFocus()
'-----------------------------------------------------------------------------------
'enables appropriate buttons
'-----------------------------------------------------------------------------------
     
    EnableButtons
    EnableOKButton

End Sub

'------------------------------------------------------------------------------------
Private Sub lstMACROModules_Click()
'------------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------------
    mbChanged = True
End Sub
'--------------------------------------------------------------------------------------------------
Private Sub AddToSelectedRoles(ByVal nNumberToAdd As Integer)
'--------------------------------------------------------------------------------------------------
'Adds role to selected role list box and removes from available role list box
'if nNumberToAdd = 0 then selected items only added
'if nNumberToAdd = 1 then all items to be added
'--------------------------------------------------------------------------------------------------
Dim sText As String
Dim i As Integer
Dim j As Integer

    On Error GoTo ErrHandler

    'Loops through the list box to check for multi-selected items then adds to list box
    If nNumberToAdd = 0 Then
        For i = 0 To lstAvailableRoles.ListCount - 1
            If Not DoesRoleExist(lstSelectedRoles, lstAvailableRoles.List(i)) Then
                If lstAvailableRoles.Selected(i) = True Then
                    'Text from the list box
                    sText = lstAvailableRoles.List(i)
                    'Add text from lstAvailableRoles to lstSelectedRoles
                    lstSelectedRoles.AddItem sText
                    lstSelectedRoles.Selected(lstSelectedRoles.NewIndex) = True
                End If
            End If
        Next
        'only delete roles that do not exist in selected roles
        For j = lstAvailableRoles.ListCount - 1 To 0 Step -1
            If Not DoesRoleExist(lstSelectedRoles, lstAvailableRoles.List(i)) Then
                If lstAvailableRoles.Selected(j) = True Then
                    'Remove the text  lstAvailableRoles
                    Call lstAvailableRoles.RemoveItem(j)
                End If
            End If
        Next
        lstAvailableRoles.ListIndex = -1
        lstAvailableRoles_Click
    End If
    
    ' loops through the list box and adds all
    If nNumberToAdd = 1 Then
        For i = 0 To lstAvailableRoles.ListCount - 1
            If Not DoesRoleExist(lstSelectedRoles, lstAvailableRoles.List(i)) Then
                'Text from the list box
                sText = lstAvailableRoles.List(i)
                'Add text from lstAvailableRoles to lstSelectedRoles
                lstSelectedRoles.AddItem sText
            End If
        Next
        'loops through the list box and deletes all items
        For j = lstAvailableRoles.ListCount - 1 To 0 Step -1
            If Not DoesRoleExist(lstAvailableRoles, lstAvailableRoles.List(i)) Then
                    'Remove the text  lstAvailableRoles
                    Call lstAvailableRoles.RemoveItem(j)
            End If
        Next
    'disable button since nothing to delete
    cmdRemoveAll.Enabled = False
    End If

    'ensures that no items in the lstAvailableRoles list box are selected
    lstAvailableRoles.ListIndex = -1

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmRoleManagement.AddToSelectedRoles"
End Sub

'--------------------------------------------------------------------------------------------------
Private Sub RemoveFromSelectedRoles(ByVal nNumberToAdd As Integer)
'--------------------------------------------------------------------------------------------------
'Adds role to available role list box and removes from selected role list box
'if nNumberToAdd = 0 then selected items only added
'if nNumberToAdd = 1 then all items to be added
'--------------------------------------------------------------------------------------------------
Dim sText As String
Dim i As Integer

    On Error GoTo ErrHandler
    
    'remove selected item(s) only
    If nNumberToAdd = 0 Then
        'Loops through the list box to check for selected items
        For i = lstSelectedRoles.ListCount - 1 To 0 Step -1
            If lstSelectedRoles.Selected(i) = True Then
                'Text from the list box
                sText = lstSelectedRoles.List(i)
                'Add text
                If AddToSelectedMACROModule(sText) = True Then
                    lstAvailableRoles.AddItem sText
                    lstAvailableRoles.Selected(lstAvailableRoles.NewIndex) = True
                End If
                'Remove the text
                Call lstSelectedRoles.RemoveItem(i)
            End If
        Next
    End If
    
    'remove all items
    If nNumberToAdd = 1 Then
        'Loops through the list box to check for multiselected items
        For i = lstSelectedRoles.ListCount - 1 To 0 Step -1
                'Text from the list box
                sText = lstSelectedRoles.List(i)
                'Add text
                If AddToSelectedMACROModule(sText) Then
                    lstAvailableRoles.AddItem sText
                End If
                'Remove the text
                Call lstSelectedRoles.RemoveItem(i)
        Next
        'disable button
        cmdRemoveAll.Enabled = False
    End If

    'ensures that no items in the selected roles listbox are selected
    lstSelectedRoles.ListIndex = -1
            
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmRoleManagement.RemoveFromSelectedRoles"
End Sub

'---------------------------------------------------------------------------------
Private Sub lstMACROModules_DblClick()
'---------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------
    mbChanged = True
    
    AddToSelectedRoles (0)

End Sub

'----------------------------------------------------------------------------------
Private Function DoesRoleExist(oList As ListBox, ByVal sRole As String) As Boolean
'----------------------------------------------------------------------------------
'checks if the role to be added is already in the list box
'----------------------------------------------------------------------------------
Dim n As Integer

    On Error GoTo ErrHandler
    
    DoesRoleExist = False
    
    For n = 0 To oList.ListCount - 1
        If oList.List(n) = sRole Then
            DoesRoleExist = True
            Exit Function
        End If
    Next

Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmRoleManagement.DoesRoleExist"
End Function

'---------------------------------------------------------------------------------
Private Sub lstMACROModules_ItemCheck(Item As Integer)
'---------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------
    
    'if the checkbox is unchecked then do nothing
    If lstMACROModules.Selected(Item) = True Then
        Call ReviseAvailableAndSelectedRoles(lstMACROModules.Text, True)
    Else
        Call ReviseAvailableAndSelectedRoles(lstMACROModules.Text, False)
    End If
    
    EnableButtons
    EnableOKButton

End Sub

'---------------------------------------------------------------------------------
Private Sub lstSelectedRoles_Click()
'---------------------------------------------------------------------------------
' NCJ 28 Feb 08 - Check user's update permission
'---------------------------------------------------------------------------------

    If lstSelectedRoles.ListIndex > -1 Then
        cmdRemoveOne.Enabled = mbCanUpdate
    Else
        cmdRemoveOne.Enabled = False
    End If

End Sub

'---------------------------------------------------------------------------------
Private Sub lstSelectedRoles_DblClick()
'---------------------------------------------------------------------------------
' Remove selected roles from list
'---------------------------------------------------------------------------------
    
    ' NCJ 28 Feb 08 - Check user's role update permission
    If Not mbCanUpdate Then Exit Sub
    
    mbChanged = True
    
    RemoveFromSelectedRoles (0)
    EnableButtons
    EnableOKButton
    
End Sub

'-----------------------------------------------------------------------------------
Private Sub EnableButtons()
'-----------------------------------------------------------------------------------
' NCJ 28 Feb 08 - Check user's update permission
'-----------------------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If lstSelectedRoles.ListCount > 0 Then
        cmdRemoveAll.Enabled = mbCanUpdate
    Else
        cmdRemoveAll.Enabled = False
    End If
    
    
    If lstAvailableRoles.ListCount > 0 Then
        cmdMoveAll.Enabled = mbCanUpdate
    Else
        cmdMoveAll.Enabled = False
    End If
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmRoleManagement.EnableButtons"
End Sub

'---------------------------------------------------------------------------------
Private Sub lstSelectedRoles_GotFocus()
'---------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------

    EnableButtons
    EnableOKButton

End Sub

'---------------------------------------------------------------------------------
Private Sub txtRoleCode_Change()
'---------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------
    
    If Not mbFormLoading Then
        If gblnValidString(txtRoleCode.Text, valAlpha + valNumeric + valSpace) Then
            mbChanged = True
            lstMACROModules.Enabled = True
            EnableOKButton
        Else
            DialogInformation "Role code cannot contain invalid characters"
            txtRoleCode.SetFocus
            txtRoleCode.Text = txtRoleCode.Tag
        End If
    End If
    
    txtRoleCode.Tag = txtRoleCode.Text
    
End Sub


'------------------------------------------------------------------------------------
Private Sub UpdateRoles()
'------------------------------------------------------------------------------------
'Updates  roles
'------------------------------------------------------------------------------------
'Dim sRoleCode As String
Dim sRoleDescription As String
Dim sFunctionCode As String
Dim rsFunctionCode As ADODB.Recordset
Dim rsRole As ADODB.Recordset
Dim sSQL As String
Dim sFunction As String
Dim n As Integer
Dim nEnabled As Integer
Dim sTaskId As String
Dim sMessage As String
Dim sFunctions As String
Dim oSystemMessage As SysMessages
Dim sMessageParameters As String
Dim nSysAdmin As Integer
Dim blExist As Boolean
Dim i As Integer

    On Error GoTo ErrHandler
    TransBegin
        
        msRoleCode = txtRoleCode.Text
        sRoleDescription = txtRoleDescription.Text
        nEnabled = 1
        nSysAdmin = chkSysAdmin.Value
        blExist = False
        
        sSQL = "SELECT * FROM Role"
        Set rsRole = New ADODB.Recordset
        rsRole.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
        
        If rsRole.RecordCount <= 0 Then Exit Sub
        
        rsRole.MoveFirst
        If mbNewRole Then
            For i = 1 To rsRole.RecordCount
                If UCase(rsRole!RoleCode) = Trim(UCase(msRoleCode)) Then
                    DialogInformation ("Rolecode " & Trim(msRoleCode) & " already exists.")
                    blExist = True
                    Exit Sub
                End If
                rsRole.MoveNext
            Next
        End If
        
        If blExist = False And mbNewRole Then
            sSQL = "INSERT INTO Role VALUES ('" & msRoleCode & "','" & sRoleDescription & "'," & nEnabled & "," & nSysAdmin & ")"
            SecurityADODBConnection.Execute sSQL, , adCmdText
        Else
            sSQL = "UPDATE Role SET RoleDescription = '" & sRoleDescription & "'," _
                & " SysAdmin = " & nSysAdmin _
                & " WHERE RoleCode = '" & msRoleCode & "'"
            SecurityADODBConnection.Execute sSQL, , adCmdText
        End If
        
        sSQL = "DELETE FROM RoleFunction WHERE RoleCode ='" & msRoleCode & "'"
        SecurityADODBConnection.Execute sSQL, , adCmdText
    
        For n = 0 To lstSelectedRoles.ListCount - 1
            sFunctionCode = lstSelectedRoles.List(n)
            sSQL = "SELECT FunctionCode FROM MACROFunction"
            sSQL = sSQL & " WHERE MACROFunction = '" & sFunctionCode & "'"
            Set rsFunctionCode = New ADODB.Recordset
            rsFunctionCode.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
            sFunction = rsFunctionCode.Fields(0).Value
        
            sSQL = "INSERT INTO RoleFunction " _
            & " VALUES ('" & msRoleCode & "','" & sFunction & "'" & ")"
            SecurityADODBConnection.Execute sSQL, , adCmdText
            
            sFunctions = sFunctions & sFunction
            If n <> lstSelectedRoles.ListCount - 1 Then
                sFunctions = sFunctions & gsSEPARATOR
            End If
        Next
        
        Set oSystemMessage = New SysMessages
        sMessageParameters = msRoleCode & gsPARAMSEPARATOR & sRoleDescription & gsPARAMSEPARATOR & nEnabled & gsPARAMSEPARATOR & nSysAdmin & gsPARAMSEPARATOR & sFunctions
        Call oSystemMessage.AddNewSystemMessage(MacroADODBConnection, ExchangeMessageType.Role, goUser.UserName, "", "Role Management", sMessageParameters)
        Set oSystemMessage = Nothing
    
    TransCommit
         
    'refresh the tree view
    Call frmMenu.RefereshTreeView
    
    If mbNewRole Then
        sTaskId = gsNEW_ROLE
        sMessage = "A new role" & " " & msRoleCode & " was created."
    Else
        sTaskId = gsEDIT_ROLE
        sMessage = "Role " & msRoleCode & " was edited."
    End If
    
    'log the creation of a new role
    Call goUser.gLog(goUser.UserName, sTaskId, sMessage)

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmRoleManagement.UpdateRoles"
End Sub

'-----------------------------------------------------------------------------------------
Private Sub GetAllRoleFunctions(ByVal sCriteria As String)
'-----------------------------------------------------------------------------------------
'gets all currently available macro role functions for all/selected macro modules
'-----------------------------------------------------------------------------------------
Dim sModule As String
    
    On Error GoTo ErrHandler
        
    Select Case sCriteria
        
        Case msMACRO_SYSTEM_MNGT
            Set mcolSystemManagement = New Collection
            Call LoadModuleCollection(msMACRO_SYSTEM_MNGT_PREFIX, mcolSystemManagement)
               
        Case msMACRO_CREATE_DATA_VIEWS
            Set mcolCreateDataViews = New Collection
            Call LoadModuleCollection(msMACRO_CREATE_DATA_VIEWS_PREFIX, mcolCreateDataViews)
        
        Case msMACRO_STUDY_DEF
            Set mcolStudyDefinition = New Collection
            Call LoadModuleCollection(msMACRO_STUDY_DEF_PREFIX, mcolStudyDefinition)
        
        Case msMACRO_LIBRARY_MNGT
            Set mcolLibraryManagement = New Collection
            Call LoadModuleCollection(msMACRO_LIBRARY_MNGT_PREFIX, mcolLibraryManagement)
        
        Case msMACRO_DATA_ENTRY
            Set mcolDataEntry = New Collection
            Call LoadModuleCollection(msMACRO_DATA_ENTRY_PREFIX, mcolDataEntry)
        
        Case msMACRO_DATA_REVIEW
            Set mcolDataReview = New Collection
            Call LoadModuleCollection(msMACRO_DATA_REVIEW_PREFIX, mcolDataReview)
        
        Case msMACRO_QUERY_MODULE
            Set mcolQueryModule = New Collection
            Call LoadModuleCollection(msMACRO_QUERY_MODULE_PREFIX, mcolQueryModule)
        
        Case msMACRO_BATCH_DATA_ENTRY
            Set mcolBatchDataEntryModule = New Collection
            Call LoadModuleCollection(msMACRO_BATCH_DATA_ENTRY_PREFIX, mcolBatchDataEntryModule)

        ' NCJ 4 Mar 03
        Case msMACRO_BATCH_VALIDATION
            Set mcolBatchValidationModule = New Collection
            Call LoadModuleCollection(msMACRO_BATCH_VALIDATION_PREFIX, mcolBatchValidationModule)

    End Select
          
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmRoleManagement.GetAllRoleFunctions"
End Sub

'-----------------------------------------------------------------------------------------------
Private Sub GetRoles()
'-----------------------------------------------------------------------------------------------
'
'-----------------------------------------------------------------------------------------------
Dim sRoleCode As String
Dim rsFunctionCode As ADODB.Recordset
Dim sSQL As String
Dim sDescription As String
Dim i As Integer
Dim tempCol As Collection
Dim n As Integer
Dim nSysAdmin As Integer
 
 
    On Error GoTo ErrHandler
    
    Set tempCol = New Collection
    
    If mbNewRole Then
        txtRoleDescription.Text = ""
        txtRoleCode.Enabled = True
        txtRoleCode.SetFocus
        txtRoleCode.Text = msRoleCode
        chkSysAdmin.Value = 0
        'clear listbox
        lstSelectedRoles.Clear
    
    Else
        'clear listbox
        lstSelectedRoles.Clear

            Set rsFunctionCode = gdsGetFnCode(msRoleCode)
            With rsFunctionCode
                Do While .EOF = False
                tempCol.Add rsFunctionCode("MACROFunction").Value, rsFunctionCode("MACROFunction").Value
                .MoveNext
            Loop
            End With
    
        'load from collection into list box
        For i = 1 To tempCol.Count
            lstSelectedRoles.AddItem tempCol.Item(i)
        Next
    
        'remove items that exist in both listboxes
        For i = tempCol.Count To 1 Step -1
            If CollectionMember(mcolRoles, tempCol.Item(i), False) Then
                mcolRoles.Remove tempCol.Item(i)
            End If
        Next
    
        'get role description for display in textbox
        sSQL = "Select RoleDescription, SysAdmin from Role Where RoleCode = '" & msRoleCode & "'"
        Set rsFunctionCode = New ADODB.Recordset
        rsFunctionCode.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

        sDescription = rsFunctionCode!roledescription
        nSysAdmin = rsFunctionCode!SysAdmin
        txtRoleDescription.Text = sDescription
        txtRoleCode.Text = msRoleCode
        txtRoleCode.Enabled = False
        chkSysAdmin.Value = nSysAdmin
        DoEvents
    
        'enable button to allow all roles to be moved
        cmdMoveAll.Enabled = True
        Set tempCol = Nothing
    End If
          
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmRoleManagement.GetRoles"
End Sub

'---------------------------------------------------------------------------------
Private Sub CheckAllModules()
'---------------------------------------------------------------------------------
'checks all the modules in the list box
'---------------------------------------------------------------------------------
Dim n As Integer

    On Error GoTo ErrHandler
    
    For n = 0 To lstMACROModules.ListCount
        Call ListCtrl_ListSelect(lstMACROModules, lstMACROModules.List(n))
        lstMACROModules.Enabled = False
    Next

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmRoleManagement.CheckAllModules"
End Sub

'-----------------------------------------------------------------------------------
Private Sub ReviseAvailableAndSelectedRoles(ByVal sModule As String, _
                                        ByVal bAdd As Boolean)
'-----------------------------------------------------------------------------------
'revises roles in the available and selected roles list boxes based on a selection
'or de-selection made in the MACROModules list box.
'sCall used to check if this routine is called from Form_Load, else this routine will
'be called the number = to current macro modules when the form loads
'-----------------------------------------------------------------------------------
Dim sSQL As String
Dim rsRoles As ADODB.Recordset
Dim sCriteria As String
Dim sRoleCode As String

    On Error GoTo ErrHandler
    
    If mbFormLoading = True Then Exit Sub
    sRoleCode = txtRoleCode.Text
    
    Select Case sModule
        Case msMACRO_SYSTEM_MNGT
            sCriteria = msMACRO_SYSTEM_MNGT_PREFIX
        Case msMACRO_STUDY_DEF
             sCriteria = msMACRO_STUDY_DEF_PREFIX
        Case msMACRO_DATA_ENTRY
             sCriteria = msMACRO_DATA_ENTRY_PREFIX
        Case msMACRO_CREATE_DATA_VIEWS
             sCriteria = msMACRO_CREATE_DATA_VIEWS_PREFIX
        Case msMACRO_QUERY_MODULE
             sCriteria = msMACRO_QUERY_MODULE_PREFIX
        Case msMACRO_DATA_REVIEW
             sCriteria = msMACRO_DATA_REVIEW_PREFIX
        Case msMACRO_LIBRARY_MNGT
             sCriteria = msMACRO_LIBRARY_MNGT_PREFIX
        Case msMACRO_BATCH_DATA_ENTRY
             sCriteria = msMACRO_BATCH_DATA_ENTRY_PREFIX
        ' NCJ 4 Mar 03
        Case msMACRO_BATCH_VALIDATION
             sCriteria = msMACRO_BATCH_VALIDATION_PREFIX
    End Select
    
    If sCriteria = "" Then Exit Sub
    
'    Set colUserRoles = New Collection
    
    sSQL = "Select MACROFunction.MACROFunction"
    sSQL = sSQL & " FROM MACROFunction,FunctionModule"
    sSQL = sSQL & " WHERE FunctionModule.MACROModule = '" & sCriteria & "'"
    sSQL = sSQL & " AND MACROFunction.FunctionCode = FunctionModule.FunctionCode"
    Set rsRoles = New ADODB.Recordset
    rsRoles.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsRoles.EOF Then Exit Sub
    
'    rsRoles.MoveFirst
'
'    'add to collection
'    Do Until rsRoles.EOF
'        colUserRoles.Add rsRoles("MACROFunction").Value, rsRoles("MACROFunction").Value
'        rsRoles.MoveNext
'    Loop
    
    If Not bAdd Then
        Call DeleteAvailableRoles(sCriteria)
    Else
        Call UpdateAvailableRoles(rsRoles)
    End If
    
    Call rsRoles.Close
    Set rsRoles = Nothing
    
'    Set colUserRoles = Nothing

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmRoleManagement.ReviseAvailableAndSelectedRoles"
End Sub

'--------------------------------------------------------------------------------
Private Function gdsGetTempFunctionCode(tmpRoleFunction As String)
'--------------------------------------------------------------------------------
' Get the FunctionCode from the FunctionName selected so we can then use the
' FunctionCode so we can complete the Insert in AddFunctionToRole
'--------------------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

    sSQL = "SELECT FunctionCode FROM Function" _
            & " WHERE Function = '" & tmpRoleFunction & "'"
            
    SecurityADODBConnection.Execute sSQL, , adCmdText
 
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmRoleManagement.gdsGetTempFunctionCode"
End Function

'----------------------------------------------------------------------------------
Private Sub DeleteAvailableRoles(ByVal sModule As String)
'----------------------------------------------------------------------------------
'
'----------------------------------------------------------------------------------
Dim n As Integer
Dim i As Integer
Dim tempCol As Collection

    On Error GoTo ErrHandler
    
    Set tempCol = New Collection
    
    'add items in list box to collection
    For n = 0 To lstAvailableRoles.ListCount - 1
        If MoreThanOneMacroModule(lstAvailableRoles.List(n)) Then
            tempCol.Add lstAvailableRoles.List(n), lstAvailableRoles.List(n)
        End If
    Next
        
    'load list box with revised items from collection
    lstAvailableRoles.Clear
    For i = 1 To tempCol.Count
        lstAvailableRoles.AddItem tempCol.Item(i)
    Next
    
    Set tempCol = Nothing

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmRoleManagement.DeleteAvailableRoles"
End Sub

'------------------------------------------------------------------------------------
Private Sub UpdateAvailableRoles(ByVal rsRoles As ADODB.Recordset)
'------------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------------
Dim n As Integer

    On Error GoTo ErrHandler

    rsRoles.MoveFirst
    n = rsRoles.RecordCount
    Do Until rsRoles.EOF
        If rsRoles.EOF Then Exit Sub
        If Not DoesRoleExist(lstAvailableRoles, rsRoles!MACROFunction) And _
            Not DoesRoleExist(lstSelectedRoles, rsRoles!MACROFunction) Then
            lstAvailableRoles.AddItem rsRoles!MACROFunction
        End If
        rsRoles.MoveNext
    Loop

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmRoleManagement.UpdateAvailableRoles"
End Sub

'---------------------------------------------------------------------------------------
Private Function AddToSelectedMACROModule(ByVal sRoleToAdd As String) As Boolean
'---------------------------------------------------------------------------------------
'finds which of the macro modules have been checked before adding items from the selected roles
'list view to the available roles list view. If the role does not belong to the selected module
'the roles are deleted from the selected roles list but not added to the available roles.
'---------------------------------------------------------------------------------------
Dim n As Integer

    On Error GoTo ErrHandler
    
    AddToSelectedMACROModule = False
    
    For n = 0 To lstMACROModules.ListCount - 1
        If lstMACROModules.Selected(n) = True Then
            Select Case lstMACROModules.List(n)
                Case msMACRO_DATA_ENTRY
                    AddToSelectedMACROModule = CollectionMember(mcolDataEntry, sRoleToAdd, False)
                Case msMACRO_DATA_REVIEW
                    AddToSelectedMACROModule = CollectionMember(mcolDataReview, sRoleToAdd, False)
                Case msMACRO_STUDY_DEF
                    AddToSelectedMACROModule = CollectionMember(mcolStudyDefinition, sRoleToAdd, False)
                Case msMACRO_QUERY_MODULE
                    AddToSelectedMACROModule = CollectionMember(mcolQueryModule, sRoleToAdd, False)
                Case msMACRO_LIBRARY_MNGT
                    AddToSelectedMACROModule = CollectionMember(mcolLibraryManagement, sRoleToAdd, False)
                Case msMACRO_CREATE_DATA_VIEWS
                    AddToSelectedMACROModule = CollectionMember(mcolCreateDataViews, sRoleToAdd, False)
                Case msMACRO_SYSTEM_MNGT
                    AddToSelectedMACROModule = CollectionMember(mcolSystemManagement, sRoleToAdd, False)
                Case msMACRO_BATCH_DATA_ENTRY
                    AddToSelectedMACROModule = CollectionMember(mcolBatchDataEntryModule, sRoleToAdd, False)
                ' NCJ 4 Mar 03
                Case msMACRO_BATCH_VALIDATION
                    AddToSelectedMACROModule = CollectionMember(mcolBatchValidationModule, sRoleToAdd, False)

            End Select
        End If
        If AddToSelectedMACROModule = True Then Exit Function
    Next


Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmRoleManagement.AddToSelectedMACROModule"
End Function

'-----------------------------------------------------------------------------
Private Sub EnableOKButton()
'-----------------------------------------------------------------------------
' NCJ 28 Feb 08 - Bug 3003 - Disable OK button if user does not have Maintain Role permission
'-----------------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
    
    ' NJC 28 Feb 08 - Bug 3003 - Check permission
    If Not mbCanUpdate Then Exit Sub
    
    ' Roles can only be applied on a server database, so leave cmdOK.enabled = false
    If GetMacroDBSetting("datatransfer", "dbtype", , gsSERVER) <> gsSERVER Then Exit Sub
    
    If mbFormLoading Then Exit Sub
    
    If lstSelectedRoles.ListCount = 0 Then Exit Sub
    
    cmdOK.Enabled = (txtRoleCode.Text <> "") And (txtRoleDescription.Text <> "")
    cmdCancel.Enabled = (txtRoleCode.Text <> "") And (txtRoleDescription.Text <> "")
    
    If cmdOK.Enabled Then
        'only system admin can apply changes to a SysAdmin Role
        If (goUser.SysAdmin = False) And (chkSysAdmin.Value = 1) Then
            cmdOK.Enabled = False
        Else
            cmdOK.Enabled = True
        End If
    End If
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmRoleManagement.EnableOKButton"
End Sub

'-------------------------------------------------------------------------------
Private Sub txtRoleDescription_Change()
'-------------------------------------------------------------------------------
    
    If Not mbFormLoading Then
        If gblnValidString(txtRoleDescription.Text, valAlpha + valNumeric + valSpace) Then
            If Len(txtRoleDescription.Text) > 255 Then
                DialogInformation "Role description may not be more than 255 characters"
                txtRoleDescription.SetFocus
                txtRoleDescription.Text = txtRoleDescription.Tag
            Else
                EnableOKButton
            End If
        Else
            DialogInformation "Role description may not contain invalid characters"
            txtRoleDescription.SetFocus
            txtRoleDescription.Text = txtRoleDescription.Tag
        End If
    End If
    
    txtRoleDescription.Tag = txtRoleDescription.Text
    
End Sub

'---------------------------------------------------------------------------------
Private Function MoreThanOneMacroModule(ByVal sRole As String) As Boolean
'---------------------------------------------------------------------------------
'this sub checks to see if an item being deleted belongs to more than one macro module.
'if it does, it will not be deleted if one of the macro module it belongs to is
'checked in the macro modules listbox
'---------------------------------------------------------------------------------
Dim n As Integer

    On Error GoTo ErrHandler
    
    MoreThanOneMacroModule = False
    
    If lstMACROModules.ListCount > 0 Then
        For n = 0 To lstMACROModules.ListCount - 1
            
            If lstMACROModules.Selected(n) = True Then
                Select Case lstMACROModules.List(n)
                    Case msMACRO_DATA_ENTRY
                        MoreThanOneMacroModule = CollectionMember(mcolDataEntry, sRole, False)
                    Case msMACRO_DATA_REVIEW
                        MoreThanOneMacroModule = CollectionMember(mcolDataReview, sRole, False)
                    Case msMACRO_STUDY_DEF
                        MoreThanOneMacroModule = CollectionMember(mcolStudyDefinition, sRole, False)
                    Case msMACRO_QUERY_MODULE
                        MoreThanOneMacroModule = CollectionMember(mcolQueryModule, sRole, False)
                    Case msMACRO_LIBRARY_MNGT
                        MoreThanOneMacroModule = CollectionMember(mcolLibraryManagement, sRole, False)
                    Case msMACRO_CREATE_DATA_VIEWS
                        MoreThanOneMacroModule = CollectionMember(mcolCreateDataViews, sRole, False)
                    Case msMACRO_SYSTEM_MNGT
                        MoreThanOneMacroModule = CollectionMember(mcolSystemManagement, sRole, False)
                    Case msMACRO_BATCH_DATA_ENTRY
                        MoreThanOneMacroModule = CollectionMember(mcolBatchDataEntryModule, sRole, False)
                    ' NCJ 4 Mar 03
                    Case msMACRO_BATCH_VALIDATION
                        MoreThanOneMacroModule = CollectionMember(mcolBatchValidationModule, sRole, False)
                End Select
            End If
            If MoreThanOneMacroModule = True Then Exit Function
        Next
    End If
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmRoleManagement.MoreThanOneMacroModule"
End Function

'--------------------------------------------------------------------------------------
Private Sub LoadModuleCollection(ByVal sModule As String, _
                                ByRef colMacroFunction As Collection)
'--------------------------------------------------------------------------------------
'sModule = MACRO Module
'colMacroFunction = collection to be loaded
'loads items into the passed MACRO module  collection
'--------------------------------------------------------------------------------------
Dim rsFunctionCode As ADODB.Recordset
Dim sSQL As String
Dim sFun As String
      
    On Error GoTo ErrHandler
  
    sSQL = "Select MACROFunction FROM MACROFunction,FunctionModule "
    sSQL = sSQL & " WHERE  MACROFunction.FunctionCode = FunctionModule.FunctionCode"
    sSQL = sSQL & " AND FunctionModule.MACROModule ='" & sModule & "'"
    Set rsFunctionCode = New ADODB.Recordset
    rsFunctionCode.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, , adCmdText

    ' NCJ 4 Mar 03 - Check for no roles
    If Not rsFunctionCode.EOF Then
        rsFunctionCode.MoveFirst
        'add macro roles to collection
        Do Until rsFunctionCode.EOF
            sFun = rsFunctionCode("MACROFunction").Value
            colMacroFunction.Add sFun, sFun
            If Not DoesRoleExist(lstAvailableRoles, sFun) And Not DoesRoleExist(lstSelectedRoles, sFun) Then
                lstAvailableRoles.AddItem sFun
            End If
            rsFunctionCode.MoveNext
        Loop
    End If
    
    Call rsFunctionCode.Close
    Set rsFunctionCode = Nothing
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmRoleManagement.LoadModuleCollection"
End Sub
