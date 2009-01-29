VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenu 
   Caption         =   "MACRO 3.0 Stub"
   ClientHeight    =   4530
   ClientLeft      =   3045
   ClientTop       =   4020
   ClientWidth     =   6360
   Icon            =   "frmMenuStub.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   6360
   Begin VB.TextBox txtMsg 
      Height          =   3195
      Left            =   300
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmMenuStub.frx":08CA
      Top             =   300
      Width           =   5775
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   435
      Left            =   4800
      TabIndex        =   0
      Top             =   3660
      Width           =   1260
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2160
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrSystemIdleTimeout 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2760
      Top             =   3720
   End
   Begin MSComctlLib.StatusBar sbrMenu 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   4155
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Key             =   "UserKey"
            Object.ToolTipText     =   "Name of current user"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Key             =   "RoleKey"
            Object.ToolTipText     =   "Role of current user"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Key             =   "UserDatabase"
            Object.ToolTipText     =   "Name of current Database"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuFHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHUserGuide 
         Caption         =   "&User Guide"
      End
      Begin VB.Menu mnuSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHAboutMacro 
         Caption         =   "&About Macro"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
' File:         frmMenuStub.frm
' Copyright:    InferMed Ltd. 2003. All Rights Reserved
' Author:       Nicky Johns, February 2003
' Purpose:      Contains the main form of the MACRO 3.0 Stub
'----------------------------------------------------------------------------------------'
'   Revisions:
'----------------------------------------------------------------------------------------'

Option Explicit

Private moArezzo As Arezzo_DM

'--------------------------------------------------------------------
Private Sub cmdExit_Click()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    ' Only shut down the ALM if it has been started
    If Not moArezzo Is Nothing Then
        moArezzo.Finish
        Set moArezzo = Nothing
    End If

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
        
    'Create and initialise a new Arezzo instance
    Set moArezzo = New Arezzo_DM
    
    ' NCJ 29 Jan 03 - Get prolog switches from new ArezzoMemory class
    Set oArezzoMemory = New clsAREZZOMemory
    Call oArezzoMemory.Load(0, goUser.CurrentDBConString)
    'Get the Prolog memory settings using GetPrologSwitches
    Call moArezzo.Init(gsTEMP_PATH, oArezzoMemory.AREZZOSwitches)
    Set oArezzoMemory = Nothing

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
    
    ' Only shut down the ALM if it has been started
    If Not moArezzo Is Nothing Then
        moArezzo.Finish
        Set moArezzo = Nothing
    End If
    
    Call ExitMACRO
    Call MACROEnd

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "Form_Unload", Err.Source) = Retry Then
        Resume
    End If
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
Public Function ForgottenPassword(sSecurityCon As String, sUserName As String, sPassword As String, sErrMsg As String) As Boolean
'---------------------------------------------------------------------
'dummy function for frmNewLogin to compile
'---------------------------------------------------------------------


End Function

