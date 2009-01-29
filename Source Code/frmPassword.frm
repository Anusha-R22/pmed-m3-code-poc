VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password Configuration"
   ClientHeight    =   2145
   ClientLeft      =   8385
   ClientTop       =   6015
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   3180
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   3015
      Begin VB.TextBox txtMinLength 
         Height          =   345
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtMaxLength 
         Height          =   345
         Left            =   1800
         TabIndex        =   1
         Top             =   660
         Width           =   1095
      End
      Begin VB.TextBox txtExpiry 
         Height          =   345
         Left            =   1800
         TabIndex        =   2
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Min password length."
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Max password length."
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Expiry period (Days)"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1140
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   540
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1860
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999. All Rights Reserved
'   File:       frmPassword.frm
'   Author:     Will Casey, October 1999
'   Purpose:    To allow the user to configure password length min and max and the
'               passwords expiry period.
'
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'Revisions:
'willC    11/10/99  Added the gLog call to log the changing of the password limits to the
'                   OK button.
'willc     10/11/99 Added the error handling
'   TA 26/04/2000   Subclassing removed
'----------------------------------------------------------------------------------------'

Option Explicit

Private nMaxPwdLength As Integer
Private nMinPwdLength As Integer
Private nPwdExpiresInDays As Long


'---------------------------------------------------------------------
Private Sub cmdCancel_Click()
'---------------------------------------------------------------------
    Unload Me
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdOK_Click()
'---------------------------------------------------------------------
' Now that you have the information regarding max and min lengths of pwds
' and the expiry period insert them into the MacroPassword table
'---------------------------------------------------------------------
'Dim nMinPwdLength As Integer
'Dim nMaxPwdLength As Integer
'Dim nPwdExpiresInDays As Long
On Error GoTo ErrHandler

    nMinPwdLength = Val(txtMinLength.Text)
    nMaxPwdLength = Val(txtMaxLength.Text)
    nPwdExpiresInDays = Val(txtExpiry.Text)
    
        
    If nMinPwdLength <= 0 Then
      MsgBox " A password must have one or more characters.", vbInformation, "MACRO"
      txtMinLength.Text = ""
      txtMinLength.SetFocus
      cmdOK.Enabled = False
      Exit Sub
    End If

    If nMaxPwdLength = 0 Then
      MsgBox " A password must have one or more characters.", vbInformation, "MACRO"
      txtMaxLength.Text = ""
      txtMaxLength.SetFocus
      cmdOK.Enabled = False
      Exit Sub
    End If

    If nMinPwdLength >= nMaxPwdLength Then
        MsgBox " The minimum length has been set as greater than or equal to the maximum length.", vbInformation, "MACRO"
        cmdOK.Enabled = False
        Exit Sub
    End If

    
    Call InsertPassWordDetails(nMinPwdLength, nMaxPwdLength, nPwdExpiresInDays)
    
    ' log the changing of the password settings
    Call gLog("sChangePasswordSettings", " The password settings were changed to min " _
            & nMinPwdLength & " chars, max " & nMaxPwdLength & " chars,expires in " & nPwdExpiresInDays & " days.")
    Unload Me
    
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "GetASEUserID")
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
Private Sub moFormIdleWatch_UserActivityDetected()
'---------------------------------------------------------------------
' restart the system idle timer
'---------------------------------------------------------------------
    Call RestartSystemIdleTimer
    
End Sub

'------------------------------------------------------------------------------'
Private Sub Form_Load()
'------------------------------------------------------------------------------'
On Error GoTo ErrHandler
    Me.BackColor = glFormColour

    Me.Icon = frmMenu.Icon

    cmdOK.Enabled = False
    
    RefreshPasswordDetails
    
    FormCentre Me
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "Form_Load")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
   
End Sub



'------------------------------------------------------------------------------'
Private Sub txtExpiry_Change()
'------------------------------------------------------------------------------'
' Set the expiry period in days
'------------------------------------------------------------------------------'
'Dim nPwdExpiresInDays As Long
On Error GoTo ErrHandler

    nPwdExpiresInDays = Val(txtExpiry.Text)
    
        If nPwdExpiresInDays > 999 Then
            MsgBox "The password expiry period must be less than 999 days.", vbInformation, "MACRO"
            txtExpiry.Text = ""
            Exit Sub
        End If

    If txtMinLength.Text = "" Or _
        txtMaxLength.Text = "" Or _
        txtExpiry.Text = "" Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtExpiry_Change")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
   
End Sub

'------------------------------------------------------------------------------'
Private Sub txtExpiry_KeyPress(KeyAscii As Integer)
'------------------------------------------------------------------------------'
' only allows numbers to be entered in the box
'------------------------------------------------------------------------------'
Dim X As String
Dim sNumbers As String
On Error GoTo ErrHandler

    X = Chr(KeyAscii)
    sNumbers = "0123456789" & vbBack & vbCr
    
    If InStr(sNumbers, X) = 0 Then
        MsgBox "Only Numeric values are allowed", vbOKOnly, "MACRO"
        KeyAscii = 8 ' backspace clears the Offending character
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtExpiry_KeyPress")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
   
End Sub

'------------------------------------------------------------------------------'
Private Sub txtMaxLength_KeyPress(KeyAscii As Integer)
'------------------------------------------------------------------------------'
' only allows numbers to be entered in the box
'------------------------------------------------------------------------------'
Dim X As String
Dim sNumbers As String
On Error GoTo ErrHandler

    X = Chr(KeyAscii)
    sNumbers = "0123456789" & vbBack & vbCr
    
    If InStr(sNumbers, X) = 0 Then
        MsgBox "Only Numeric values are allowed", vbOKOnly, "MACRO"
        KeyAscii = 8 ' backspace clears the Offending character
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtMaxLength_KeyPress")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
   
End Sub

'------------------------------------------------------------------------------'
Private Sub txtMaxLength_LostFocus()
'------------------------------------------------------------------------------'
' Check that the max length is greater than the min length
'------------------------------------------------------------------------------'
'Dim nMaxPwdLength As Integer
'Dim nMinPwdLength As Integer
On Error GoTo ErrHandler


    nMaxPwdLength = Val(txtMaxLength.Text)
    nMinPwdLength = Val(txtMinLength.Text)
    
    If nMinPwdLength > nMaxPwdLength Then
        MsgBox " The minimum length has been set as greater than the maximum length.", vbInformation, "MACRO"
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtMaxLength_LostFocus")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
   
End Sub

'------------------------------------------------------------------------------'
Private Sub txtMinLength_Change()
'------------------------------------------------------------------------------'
' Set the password min length as an integer ie up to 32,767
'------------------------------------------------------------------------------'

'Dim nMinPwdLength As Integer
On Error GoTo ErrHandler

    nMinPwdLength = Val(txtMinLength.Text)
    
     If nMinPwdLength < 1 Then
      'MsgBox " A password must have one or more characters.", vbInformation, "MACRO"
       txtMinLength.Text = ""
       'txtMinLength.SetFocus
        cmdOK.Enabled = False
    End If
    
    If nMinPwdLength > 255 Then
        MsgBox "Please choose a password length of less than 255.", vbInformation, "MACRO"
        txtMinLength.Text = ""
        cmdOK.Enabled = False
    End If
    
    If txtMinLength.Text = "" Or _
        txtMaxLength.Text = "" Or _
        txtExpiry.Text = "" Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtMinLength_Change")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
   
End Sub

'------------------------------------------------------------------------------'
Private Sub txtMaxLength_Change()
'------------------------------------------------------------------------------'
' Set the password max length as an integer ie up to 32767
'------------------------------------------------------------------------------'
'Dim nMaxPwdLength As Integer
On Error GoTo ErrHandler

    nMaxPwdLength = Val(txtMaxLength.Text)
    
    If nMaxPwdLength = 0 Then
    cmdOK.Enabled = False
        'MsgBox " A password must have one or more characters.", vbInformation, "MACRO"
        'txtMinLength.Text = ""
        txtMaxLength.SetFocus
    End If
         
    If nMaxPwdLength > 255 Then
            MsgBox "Please choose a password length of less than 255.", vbInformation, "MACRO"
         txtMaxLength.Text = ""
         cmdOK.Enabled = False
    End If

    If txtMinLength.Text = "" Or _
        txtMaxLength.Text = "" Or _
        txtExpiry.Text = "" Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtMaxLength_Change")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
   
End Sub

'------------------------------------------------------------------------------'
Private Sub txtMinLength_KeyPress(KeyAscii As Integer)
'------------------------------------------------------------------------------'
' only allows numbers to be entered in the box
'------------------------------------------------------------------------------'
Dim X As String
Dim sNumbers As String
On Error GoTo ErrHandler

    X = Chr(KeyAscii)
    sNumbers = "0123456789" & vbBack & vbCr
    
    If InStr(sNumbers, X) = 0 Then
        MsgBox "Only Numeric values are allowed", vbOKOnly, "MACRO"
        KeyAscii = 8 ' backspace clears the Offending character
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtMinLength_KeyPress")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
   
End Sub

'------------------------------------------------------------------------------'
Private Sub InsertPassWordDetails(nMinPwdLength As Integer, nMaxPwdLength As Integer, _
                                    nPwdExpiresInDays As Long)
'------------------------------------------------------------------------------'
' Insert the details for the Passwords if a record exists already do an update
' otherwise do an insert..
'------------------------------------------------------------------------------'
Dim sSQL As String
Dim rsPasswords As ADODB.Recordset
Dim nCount As Integer
On Error GoTo ErrHandler

    sSQL = "SELECT * FROM MacroPassword"
    Set rsPasswords = New ADODB.Recordset
    rsPasswords.Open sSQL, SecurityADODBConnection, adOpenKeyset, , adCmdText
    
    With rsPasswords
        nCount = rsPasswords.RecordCount
    End With
    Set rsPasswords = Nothing
    
    If nCount > 0 Then
        sSQL = "UPDATE MacroPassword SET " _
            & " MinLength = " & nMinPwdLength & "," _
            & " MaxLength = " & nMaxPwdLength & "," _
            & " ExpiryPeriod = " & nPwdExpiresInDays
             SecurityADODBConnection.Execute sSQL
    Else
        sSQL = "INSERT INTO MacroPassword" _
            & "(MinLength,MaxLength,ExpiryPeriod)" _
            & " VALUES (" & nMinPwdLength & "," & nMaxPwdLength & "," & nPwdExpiresInDays & ")"
            SecurityADODBConnection.Execute sSQL
   End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "InsertPassWordDetails")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
   
End Sub

'------------------------------------------------------------------------------'
Private Sub RefreshPasswordDetails()
'------------------------------------------------------------------------------'
' Insert the details for the Passwords on form load up
'------------------------------------------------------------------------------'

Dim sSQL As String
Dim rsPasswords As ADODB.Recordset
On Error GoTo ErrHandler

    sSQL = "SELECT * FROM MacroPassword"
    Set rsPasswords = New ADODB.Recordset
    rsPasswords.Open sSQL, SecurityADODBConnection, adOpenKeyset, , adCmdText
    
   With rsPasswords
       txtMaxLength.Text = rsPasswords!MaxLength
       txtMinLength.Text = rsPasswords!MinLength
       txtExpiry.Text = rsPasswords!ExpiryPeriod
    End With
    
    Set rsPasswords = Nothing
    
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "RefreshPasswordDetails")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
   
End Sub
