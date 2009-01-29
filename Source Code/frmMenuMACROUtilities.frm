VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MACRO Utilities"
   ClientHeight    =   2655
   ClientLeft      =   3030
   ClientTop       =   4005
   ClientWidth     =   8355
   Icon            =   "frmMenuMACROUtilities.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   8355
   Begin VB.Frame Frame3 
      Height          =   2235
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   8235
      Begin VB.Frame Frame1 
         Caption         =   "Import/Export Subject Data"
         Height          =   915
         Left            =   1200
         TabIndex        =   4
         Top             =   720
         Width           =   3255
         Begin VB.CommandButton cmdImport 
            Caption         =   "Import..."
            Height          =   345
            Left            =   300
            TabIndex        =   6
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton cmdExport 
            Caption         =   "Export..."
            Height          =   345
            Left            =   1740
            TabIndex        =   5
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Subject-Site Transfer"
         Height          =   915
         Left            =   4680
         TabIndex        =   2
         Top             =   720
         Width           =   2595
         Begin VB.CommandButton cmdTransSubject 
            Caption         =   "Transfer Subject..."
            Height          =   345
            Left            =   300
            TabIndex        =   3
            Top             =   360
            Width           =   1995
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1440
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrSystemIdleTimeout 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2100
      Top             =   1680
   End
   Begin MSComctlLib.StatusBar sbrMenu 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Key             =   "UserKey"
            Object.ToolTipText     =   "Name of current user"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Text            =   "Role"
            TextSave        =   "Role"
            Key             =   "RoleKey"
            Object.ToolTipText     =   "Role of current user"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Key             =   "UserDatabase"
            Object.ToolTipText     =   "Name of current Database"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Bevel           =   0
            TextSave        =   "28/10/2003"
            Object.ToolTipText     =   "Date"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Bevel           =   0
            TextSave        =   "12:54"
            Object.ToolTipText     =   "Time"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuDataExport 
      Caption         =   "&Data"
      Begin VB.Menu mnuSubject 
         Caption         =   "&Subject"
         Begin VB.Menu mnuExport 
            Caption         =   "&Export..."
         End
         Begin VB.Menu mnuImport 
            Caption         =   "&Import..."
         End
         Begin VB.Menu mnuTransferSubject 
            Caption         =   "Subject-Site Transfer..."
         End
      End
   End
   Begin VB.Menu mnuFHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHAboutMacro 
         Caption         =   "&About MACRO"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
' File:         frmMenuMACROUtilities.frm
' Copyright:    InferMed Ltd. 2003. All Rights Reserved
' Author:       Richard Meinesz September 2003
' Purpose:      Main menu form of the MACRO 3.0 Utilities module
'----------------------------------------------------------------------------------------'
' Revisions:
'   NCJ 28 Oct 03 - Changed caption to "MACRO Utilities"; added permission check in InitialiseMe
'----------------------------------------------------------------------------------------'

Option Explicit

'--------------------------------------------------------------------
Private Sub cmdExit_Click()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler

    Call ExitMACRO
    Call MACROEnd

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "cmdExit_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Public Sub InitialiseMe()
'--------------------------------------------------------------------
Dim oArezzoMemory As clsAREZZOMemory

    On Error GoTo ErrHandler
    
    'The following Doevents prevents command buttons ghosting during form load
    DoEvents

    ' NCJ 28 Oct 03 - Check System Mgmnt permission here
    ' (only temporary because eventually these routines will be
    ' inside System Management itself)
    If Not goUser.CheckPermission(gsFnSystemManagement) Then
        ' Throw out unauthorised intruders
        Call DialogError("You do not have permission to access the MACRO Utilities module")
        Call cmdExit_Click
    End If
    
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "InitialiseMe", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Public Sub CheckUserRights()
'--------------------------------------------------------------------
' Dummy routine which gets called during MACRO initialisation
'--------------------------------------------------------------------

End Sub

'--------------------------------------------------------------------
Private Sub cmdExport_Click()
'--------------------------------------------------------------------

    Call frmExportPatientData.Display

End Sub

'--------------------------------------------------------------------
Private Sub cmdImport_Click()
'--------------------------------------------------------------------

    Call DisplayImportPatientDataForm

End Sub

'--------------------------------------------------------------------
Private Sub DisplayImportPatientDataForm()
'--------------------------------------------------------------------

    frmImportPatientData.Display (goUser.Database.DatabaseCode)

End Sub

'--------------------------------------------------------------------
Private Sub cmdTransSubject_Click()
'--------------------------------------------------------------------
'REM 17/09/03
'
'--------------------------------------------------------------------

    Call frmSiteTransfer.Display

End Sub

'--------------------------------------------------------------------
Private Sub Form_Load()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    FormCentre Me

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "Form_Load", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Private Sub Form_Resize()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    'check that form has not been minimised
    If Me.WindowState = vbMinimized Then Exit Sub
    
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "Form_Resize", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    Call ExitMACRO
    Call MACROEnd

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "Form_Unload", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Private Sub mnuExport_Click()
'--------------------------------------------------------------------

    Call frmExportPatientData.Display

End Sub

'--------------------------------------------------------------------
Private Sub mnuFExit_Click()
'--------------------------------------------------------------------

    Call cmdExit_Click

End Sub

'--------------------------------------------------------------------
Private Sub mnuHAboutMacro_Click()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    frmAbout.Display

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "mnuHAboutMacro_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------
Private Sub mnuHUserGuide_Click()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    Call MACROHelp(Me.hWnd, App.Title)

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "mnuHUserGuide_Click", Err.Source) = Retry Then
        Resume
    End If
End Sub

'---------------------------------------------------------------------
Public Function ForgottenPassword(sSecurityCon As String, sUsername As String, sPassword As String, sErrMsg As String) As Boolean
'---------------------------------------------------------------------
'dummy function for frmNewLogin to compile
'---------------------------------------------------------------------


End Function

'---------------------------------------------------------------------
Private Sub mnuImport_Click()
'---------------------------------------------------------------------

    Call DisplayImportPatientDataForm

End Sub

'---------------------------------------------------------------------
Private Sub mnuTransferSubject_Click()
'---------------------------------------------------------------------

    Call frmSiteTransfer.Display
    
End Sub
