VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserRoleErrors 
   Caption         =   "User Role Conflicts"
   ClientHeight    =   4215
   ClientLeft      =   9645
   ClientTop       =   4560
   ClientWidth     =   6075
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6075
   Begin VB.CommandButton cmdNo 
      Caption         =   "No"
      Height          =   325
      Left            =   3060
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes"
      Height          =   325
      Left            =   1680
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Frame fraUserRole 
      Caption         =   "User Role Conflicts"
      Height          =   2295
      Left            =   60
      TabIndex        =   1
      Top             =   1020
      Width           =   5955
      Begin MSComctlLib.ListView lvwRoleErrors 
         Height          =   1515
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   2672
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblRoles 
         Alignment       =   2  'Center
         Caption         =   "The following user roles conflict with your selection and will be deleted"
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   300
         Width           =   5175
      End
   End
   Begin VB.Frame fraTrial 
      Caption         =   "Study Site  Messages"
      Height          =   860
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5955
      Begin VB.Label lblTrialSiteMsg 
         Alignment       =   2  'Center
         Caption         =   "The user role you have selected will have no effect because "
         Height          =   430
         Left            =   720
         TabIndex        =   7
         Top             =   260
         Width           =   4395
      End
   End
   Begin VB.Label lblWaring 
      Alignment       =   2  'Center
      Caption         =   "Are you sure you want to create the specified user role ?"
      Height          =   315
      Left            =   840
      TabIndex        =   5
      Top             =   3420
      Width           =   4395
   End
End
Attribute VB_Name = "frmUserRoleErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   File:       frmUserRoleErrors.frm
'   Author:     Ashitei Trebi-Ollennu, October 2002
'   Purpose:    Used to replace message box showing conflicting user roles when user roles
'               are being added in form frmNWUserRole
'------------------------------------------------------------------------------

Option Explicit
Private mbOKPressed As Boolean

'-------------------------------------------------------------------
Public Function Display(ByVal sRoleConflicts As String, _
                        ByVal sTrialMsg As String) As Boolean
'-------------------------------------------------------------------
'
'-------------------------------------------------------------------
Dim var As Variant
Dim vAny As Variant
Dim itmX As MSComctlLib.ListItem
Dim n As Integer

    On Error GoTo ErrHandler

    
    Me.Icon = frmMenu.Icon
    FormCentre Me
    
    mbOKPressed = False
    
    ShowColumnHeaders
       
    If sRoleConflicts <> "" Then
        
        sRoleConflicts = Right(sRoleConflicts, Len(sRoleConflicts) - Len(vbCrLf))
        var = Split(sRoleConflicts, vbCrLf)
        For n = 0 To UBound(var)
            vAny = Split(var(n), "|")
            Set itmX = lvwRoleErrors.ListItems.Add(, , vAny(0))
            itmX.SubItems(1) = vAny(1)
            itmX.SubItems(2) = vAny(2)
        Next
        Call Resize(True)
    ElseIf sTrialMsg <> "" Then
        lblTrialSiteMsg.Caption = "The user role you have selected will have no effect because " & vbCrLf & sTrialMsg
        Call Resize(False)
    End If
    
    Call lvw_SetAllColWidths(lvwRoleErrors, LVSCW_AUTOSIZE_USEHEADER)
    
    HourglassSuspend
    Me.Show vbModal
    HourglassResume
    
    Display = mbOKPressed
     
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmUserRoleErrors.Display"
End Function

'--------------------------------------------------------------
Private Sub cmdNo_Click()
'--------------------------------------------------------------
'
'--------------------------------------------------------------
    
    Unload Me

End Sub

'-----------------------------------------------------------------
Private Sub cmdYes_Click()
'-----------------------------------------------------------------
'
'-----------------------------------------------------------------
    
    mbOKPressed = True
    Unload Me

End Sub

'-------------------------------------------------------------------
Private Sub Form_Load()
'-------------------------------------------------------------------
'
'-------------------------------------------------------------------
    
    FormCentre Me
    Me.Icon = frmMenu.Icon

End Sub
'------------------------------------------------------------------------------------
Private Sub ShowColumnHeaders()
'------------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------------
Dim colmX As MSComctlLib.ColumnHeader

    'add column headers with widths
    Set colmX = lvwRoleErrors.ColumnHeaders.Add(, , "Study", 1700)
    Set colmX = lvwRoleErrors.ColumnHeaders.Add(, , "Site", 1700)
    Set colmX = lvwRoleErrors.ColumnHeaders.Add(, , "Role", 1700)
     
    'set view type
    lvwRoleErrors.View = lvwReport
    'set initial sort to ascending on column 0 (Study)
    lvwRoleErrors.SortKey = 0
    lvwRoleErrors.SortOrder = lvwAscending

End Sub

'----------------------------------------------------------------------------------------'
Private Sub Resize(ByVal bShowConflicts As Boolean)
'----------------------------------------------------------------------------------------'
'REM 29/10/02
'Resize form according to which message is being displayed
'----------------------------------------------------------------------------------------'
Dim lTitleHeight As Long


    lTitleHeight = Me.Height - Me.ScaleHeight
    
    fraUserRole.Visible = bShowConflicts
    lblRoles.Visible = bShowConflicts
    lvwRoleErrors.Visible = bShowConflicts
    
    lblWaring.Visible = True
    
    fraTrial.Visible = Not bShowConflicts
    lblTrialSiteMsg.Visible = Not bShowConflicts
    
    If bShowConflicts Then
        fraUserRole.Top = 120
        lblWaring.Top = 2500
        cmdNo.Top = fraUserRole.Top + fraUserRole.Height + 430
    Else
        lblWaring.Top = 1020
        cmdNo.Top = fraTrial.Top + fraTrial.Height + 390
    End If
    
    cmdYes.Top = cmdNo.Top
    
    Me.Height = lTitleHeight + cmdNo.Top + cmdNo.Height + 60
    
End Sub
