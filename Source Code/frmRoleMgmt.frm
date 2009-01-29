VERSION 5.00
Begin VB.Form frmRoleMgmt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Role Management"
   ClientHeight    =   3330
   ClientLeft      =   4485
   ClientTop       =   4170
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   2340
      Width           =   7215
      Begin VB.CommandButton cmdEditRole 
         Caption         =   "&Edit Role"
         Height          =   495
         Left            =   2040
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdNewRole 
         Caption         =   "&Add New Role"
         Height          =   495
         Left            =   3780
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Close"
         Height          =   495
         Left            =   5520
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.ListBox lstUserFunction 
      Height          =   1425
      Left            =   3840
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   3525
   End
   Begin VB.ListBox lstUserRole 
      Height          =   1425
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   3525
   End
   Begin VB.Label lblDescription 
      Height          =   435
      Left            =   120
      TabIndex        =   8
      Top             =   1980
      Width           =   7215
   End
   Begin VB.Label Label2 
      Caption         =   "Role functions:"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Available roles:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmRoleMgmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       frmRoleMgmt.frm
'   Author:     Will Casey 20/10/99
'   Purpose:    Maintain Roles and Functions.
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
'   Revisions:
'willC      11/10/99 Added the error handlers
'   TA 08/05/2000   removed subclassing
'WillC SR3756 31/8/00 Added label to show role description
'--------------------------------------------------------------------------------

Option Explicit
Option Compare Binary
Option Base 0


'----------------------------------------------------------------------------------------------
Private Sub cmdEditRole_Click()
'----------------------------------------------------------------------------------------------
' Show the form to edit a user role.
'----------------------------------------------------------------------------------------------
On Error GoTo ErrHandler
    
    frmEditUserRole.Show vbModal
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdEditRole_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub


'----------------------------------------------------------------------------------------------
Private Sub cmdNewRole_Click()
'----------------------------------------------------------------------------------------------
' The RoleCode is the user friendly name for a Role ie Monitor,Investigator
'----------------------------------------------------------------------------------------------
    
    frmNewUserRole.Show vbModal

End Sub

'----------------------------------------------------------------------------------------------
Private Sub Form_Load()
'----------------------------------------------------------------------------------------------
'Open the Database, and Load all the Lists with the data, add the columns to the
' listview, Disable any buttons that could cause a problem at this stage.
'----------------------------------------------------------------------------------------------
On Error GoTo ErrHandler

    Me.Icon = frmMenu.Icon
    
' Turn on key preview for form, so that F1 (Help) can be trapped by form
    Me.KeyPreview = True
    
    If goUser.CheckPermission(gsFnMaintRole) = False Then
        cmdEditRole.Enabled = False
    End If
    
    
' Disable the EditRole button until a role has been chosen for editing
    cmdEditRole.Enabled = False
    
    ' Get the latest data from the database
    RefreshRoles
    RefreshFunctions
    
    
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

'----------------------------------------------------------------------------------------------
Private Sub lstUserRole_Click()
'----------------------------------------------------------------------------------------------
' when a Role is clicked show the relevant functions for that role.
'----------------------------------------------------------------------------------------------
 Dim sRoleCode As String
 Dim rsFunctionCode As ADODB.Recordset
 Dim sSQL As String
 Dim sDescription As String
 
 
    On Error GoTo ErrHandler
    
    lstUserFunction.Clear
    cmdEditRole.Enabled = True

 ' Use the RoleFunction.FunctionCode to tie to the Function.FunctionCode to
 ' get the corresponding Functions for a Role
 
     sRoleCode = lstUserRole.Text
     
    Set rsFunctionCode = gdsGetFnCode(sRoleCode)
        With rsFunctionCode
           Do While .EOF = False
                'this places the role codes in the listbox
                lstUserFunction.AddItem .Fields(0).Value
                    .MoveNext
            Loop
        End With
          
    'WillC SR3756 31/8/00
    sSQL = "Select RoleDescription from Role Where RoleCode = '" & sRoleCode & "'"
    Set rsFunctionCode = New ADODB.Recordset
    rsFunctionCode.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not rsFunctionCode.EOF Then
        sDescription = rsFunctionCode!roledescription
       lblDescription.Caption = "Role description for " & sRoleCode & ": " & sDescription
        DoEvents
    End If
          
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "lstUserRole_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'----------------------------------------------------------------------------------------------
Private Sub RefreshFunctions()
'----------------------------------------------------------------------------------------------
' Add the list of Functions to the listbox
'----------------------------------------------------------------------------------------------
Dim rsFunctions As ADODB.Recordset
On Error GoTo ErrHandler

    lstUserFunction.Clear
      
    
    ' Get the list of functions where the role code selected in lstUserRole
    ' is the same as that in rolefunctions table
    Set rsFunctions = gdsFunctionList
    
    With rsFunctions
        Do Until .EOF = True
                'this places the role codes in the listbox
                lstUserFunction.AddItem .Fields(1).Value
                 .MoveNext
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

'----------------------------------------------------------------------------------------------
Private Sub RefreshRoles()
'----------------------------------------------------------------------------------------------
 ' Add the roles to their listbox
'----------------------------------------------------------------------------------------------
Dim rsRoles As ADODB.Recordset
On Error GoTo ErrHandler

     lstUserRole.Clear
     
        ' Get the list of roles
    Set rsRoles = gdsRoleList
    With rsRoles
       Do Until .EOF = True
         'this adds the RoleCode field from the table to the list box
         lstUserRole.AddItem .Fields(0).Value
                .MoveNext
       Loop
    End With
        
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

'----------------------------------------------------------------------------------------------
Private Sub cmdCancel_Click()
'----------------------------------------------------------------------------------------------
    Unload Me
    
End Sub

'----------------------------------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------------------------
' Enable the help if F1 is pushed
'----------------------------------------------------------------------------------------------
If KeyCode = vbKeyF1 Then               ' Show user guide
    'ShowDocument Me.hWnd, gsMACROUserGuidePath
    
    'REM 07/12/01 - New Call to MACRO Help
    Call MACROHelp(Me.hWnd, App.Title)
    
End If
    
End Sub

