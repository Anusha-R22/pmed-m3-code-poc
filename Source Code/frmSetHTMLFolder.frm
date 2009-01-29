VERSION 5.00
Begin VB.Form frmSetHTMLFolder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HTML Folder"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5640
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Frame fraSecureLocation 
      Caption         =   "(Optional) Secure Path"
      Height          =   1035
      Left            =   60
      TabIndex        =   7
      Top             =   1200
      Width           =   6855
      Begin VB.CommandButton cmdSecureBrowse 
         Caption         =   "&Browse..."
         Height          =   375
         Left            =   5520
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdSecureDefault 
         Caption         =   " &Default"
         Height          =   375
         Left            =   4200
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtHTMLSecureLocation 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6615
      End
   End
   Begin VB.Frame fraLocation 
      Caption         =   "Path"
      Height          =   1035
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   6855
      Begin VB.CommandButton cmdDefault 
         Caption         =   " &Default"
         Height          =   375
         Left            =   4200
         TabIndex        =   0
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse..."
         Height          =   375
         Left            =   5520
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtHTMLLocation 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6615
      End
   End
End
Attribute VB_Name = "frmSetHTMLFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   File:       frmSetHTMLFolder.frm
'   Author:     Steve Morris, Jan 2000
'   Purpose:    To allow the user to choose a folder on the hard drive
'------------------------------------------------------------------------------------'
'Revisions:
'ASH 16/04/2002 Modified routines HTMLPath,added extra parameter. Added cmdSecureBrowse,
'               cmdSecureDefault,HTMLSecureLocation_Change. Added error trappings
'------------------------------------------------------------------------------------'
Option Explicit
Private msHTMLLocation As String
Private msSecureHTMLLocation As String
Private mbOKClicked As Boolean

'------------------------------------------------------------------------------------'
Public Function HTMLPath(sPath As String, Optional sSecurePath As String) As Boolean
'------------------------------------------------------------------------------------'
'Called from outside and returns the chosen folder
'ASH 15/04/2002 Added extra optional parameter
'REM 20/01/03 - re-orgainsed form
'REM 02/04/03 - Set text box in function
'------------------------------------------------------------------------------------'
    'commented out 9/11/2001. new browser control
    'added there/4 no need for default. Ash
    'Call cmdDefault_Click
    On Error GoTo ErrHandler
    
    If sPath = "" Then
        msHTMLLocation = AddBackSlash(gsAppPath & "HTML")
        msSecureHTMLLocation = AddBackSlash(gsAppPath & "HTML")
    Else
        msHTMLLocation = sPath
        msSecureHTMLLocation = sSecurePath
    End If
    
    'REM 02/04/03 - set text boxes
    txtHTMLLocation.Text = msHTMLLocation
    txtHTMLSecureLocation.Text = msSecureHTMLLocation

    mbOKClicked = False
    
    Me.Show vbModal
    
    If mbOKClicked Then
    
        sPath = msHTMLLocation
        sSecurePath = msSecureHTMLLocation

    End If
    
    HTMLPath = mbOKClicked
   
Exit Function:
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "HTMLPath")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Function

'----------------------------------------------------------------------
Private Sub cmdBrowse_Click()
'----------------------------------------------------------------------
'Calls APIs to show Explorer like browser
'---------------------------------------------------------------------
Dim nPath As Long
Dim sPath As String
    
    On Error GoTo ErrHandler
    
    'REM 02/04/03 - Disable form so that it user has to close BrowseForFolderPath dialog to continue
    frmSetHTMLFolder.Enabled = False
        
    If msHTMLLocation <> "" Then
        sPath = BrowseForFolderByPath(msHTMLLocation)

            If sPath <> "" Then
                txtHTMLLocation.Text = AddBackSlash(sPath)
                'cmdDefault.Enabled = True
            Else
                txtHTMLLocation.Text = AddBackSlash(msHTMLLocation)
                'cmdDefault.Enabled = False
            End If
            
    Else
        msHTMLLocation = BrowseForFolder _
        (Me.hWnd, "Please select a folder for publishing files")
        txtHTMLLocation.Text = AddBackSlash(msHTMLLocation)
    End If

    frmSetHTMLFolder.Enabled = True

Exit Sub:
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdBrowse_Click")
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
Private Function AddBackSlash(sPath As String) As String
'------------------------------------------------------------------------------------'
'REM 02/04/03
'Add a back slash onto the end of a path if it doen't have one
'------------------------------------------------------------------------------------'

    If Right(sPath, 1) <> "\" Then
        sPath = sPath & "\"
    End If
    
    AddBackSlash = sPath

End Function

Private Sub cmdCancel_Click()
'------------------------------------------------------------------------------------'
'Ash 9/11/2001 Added cancel button
'------------------------------------------------------------------------------------'
    
    Unload Me

End Sub

'------------------------------------------------------------------------------------'
Private Sub cmdDefault_Click()
'------------------------------------------------------------------------------------'
'   Sets the folder to the default one
'------------------------------------------------------------------------------------'
    
    On Error GoTo ErrHandler
    'REM 02/04/03 - Set default regardless of what is currently in the text box
    'If Len(msHTMLLocation) = 0 Then
        msHTMLLocation = AddBackSlash(gsAppPath & "HTML\")
    'End If
    'ash 15/11/01
    'to reset default when cancel clicked on browser control
    txtHTMLLocation.Text = msHTMLLocation
    'cmdDefault.Enabled = False

Exit Sub:
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdDefault_Click")
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
Private Sub cmdOK_Click()
'------------------------------------------------------------------------------------'
'   Checks a folder has been chosen and unloads
'
'------------------------------------------------------------------------------------'
Dim sMsg As String

    On Error GoTo ErrHandler
    
    If Len(msHTMLLocation) > 0 Then
        Call goUser.Database.Load(SecurityADODBConnection, goUser.UserName, goUser.DatabaseCode, "", False, sMsg)
        mbOKClicked = True
        Unload Me
    Else
        DialogInformation "You must enter an HTML location", "HTML Folder Location"
    End If

Exit Sub:
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdOK_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'-----------------------------------------------------------------------------------'
Private Sub cmdSecureBrowse_Click()
'-----------------------------------------------------------------------------------'
'ASH Added 15/04/2002
'Calls APIs to show Explorer like browser
'-----------------------------------------------------------------------------------'
Dim nPath As Long
Dim sPath As String
    
    On Error GoTo ErrHandler
    
    'REM 02/04/03 - Disable form so that it user has to close BrowseForFolderPath dialog to continue
    frmSetHTMLFolder.Enabled = False
    
    If msSecureHTMLLocation <> "" Then
        sPath = BrowseForFolderByPath(msSecureHTMLLocation)
            If sPath <> "" Then
                txtHTMLSecureLocation.Text = AddBackSlash(sPath)
                'cmdSecureDefault.Enabled = True
            Else
                txtHTMLSecureLocation.Text = AddBackSlash(msSecureHTMLLocation)
                'cmdSecureDefault.Enabled = False
            End If
    Else
        msSecureHTMLLocation = BrowseForFolder _
        (Me.hWnd, "Please select a folder for publishing files")
        txtHTMLSecureLocation.Text = AddBackSlash(msSecureHTMLLocation)
    End If

    frmSetHTMLFolder.Enabled = True
        
Exit Sub:
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdSecureBrowse_Click")
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
Private Sub cmdSecureDefault_Click()
'------------------------------------------------------------------------------------'
'   ASH Added 15/04/2002
'   Sets the folder to the default one
'------------------------------------------------------------------------------------'
    
    On Error GoTo ErrHandler
    'REM 02/04/03 - Set default regardless of what is currently in the text box
    'If Len(msSecureHTMLLocation) = 0 Then
        msSecureHTMLLocation = gsAppPath & "HTML\"
    'End If
    'ash 15/11/01
    'to reset default when cancel clicked on browser control
    txtHTMLSecureLocation.Text = msSecureHTMLLocation
    'cmdSecureDefault.Enabled = False

Exit Sub:
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdSecureDefault_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub
'------------------------------------------------------------------------------------
Private Sub Form_Load()
'------------------------------------------------------------------------------------
'ASH 16/04/2002
'------------------------------------------------------------------------------------
    
    Me.Icon = frmMenu.Icon
    
    cmdOK.Enabled = False
'    Call cmdDefault_Click

End Sub

'------------------------------------------------------------------------------------'
Private Sub txtHTMLLocation_Change()
'------------------------------------------------------------------------------------'
'   Updates the module variable with the folder in the text box
'------------------------------------------------------------------------------------'
    On Error GoTo ErrHandler
    
    msHTMLLocation = txtHTMLLocation.Text
    txtHTMLLocation.ToolTipText = txtHTMLLocation.Text
    cmdOK.Enabled = True

Exit Sub:
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtHTMLLocation_Change")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub
'------------------------------------------------------------------------------------
Private Sub txtHTMLSecureLocation_Change()
'------------------------------------------------------------------------------------
'   Updates the module variable with the folder in the text box
'------------------------------------------------------------------------------------
'   ASH Added 15/04/2002
'------------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    
    msSecureHTMLLocation = txtHTMLSecureLocation.Text
    txtHTMLSecureLocation.ToolTipText = txtHTMLSecureLocation.Text

Exit Sub:
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtHTMLSecureLocation_Change")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub
