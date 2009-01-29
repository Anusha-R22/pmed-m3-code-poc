VERSION 5.00
Begin VB.Form frmResetPwd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reset Password"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   " Users "
      Height          =   2600
      Left            =   60
      TabIndex        =   3
      Top             =   120
      Width           =   3495
      Begin VB.ListBox lstUsers 
         Height          =   1425
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "List of known MACRO user codes"
         Top             =   240
         Width           =   3270
      End
      Begin VB.TextBox txtUserName 
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "Text1"
         ToolTipText     =   "User's full name"
         Top             =   2060
         Width           =   3270
      End
      Begin VB.Label Label1 
         Caption         =   "User Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdResetPwd 
      Caption         =   "&Reset Password"
      Default         =   -1  'True
      Height          =   375
      Left            =   60
      TabIndex        =   0
      ToolTipText     =   "Reset user's password"
      Top             =   2820
      Width           =   1400
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      ToolTipText     =   "Close this window"
      Top             =   2820
      Width           =   1400
   End
End
Attribute VB_Name = "frmResetPwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   File:       frmResetPwd.frm
'   Author:     Nicky Johns, July 2000
'   Purpose:    Form for resetting another user's password
'               in Password Management module.
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
'   Revisions:
'   NCJ 13/7/00 - Initial development (SR 3658)
'   NCJ 17/7/00 - Added gLog message when resetting password; added tooltip text to buttons
'   NCJ 18/7/00 - Changed form layout (TA's review comments)
'   Mo 18/7/2001  Changes stemming from field Password in table MacroUser being changed to
'               UserPassword (stems from the swith to Jet 4.0)
'--------------------------------------------------------------------------------

' Store the currently selected user code
Private msUserCode As String
' Store collection of user names
Private mcolUserNames As Collection

Option Explicit

'--------------------------------------------------------------------------------
Private Sub cmdClose_Click()
'--------------------------------------------------------------------------------
' CLose ourselves down
'--------------------------------------------------------------------------------

    Set mcolUserNames = Nothing
    Unload Me

End Sub

'--------------------------------------------------------------------------------
Private Sub cmdResetPwd_Click()
'--------------------------------------------------------------------------------
' Reset password of currently selected user (stored in msUserCode)
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim sMsg As String
Dim sLoginDate As String

    On Error GoTo ErrHandler
    
    If msUserCode = "" Then Exit Sub
    
    ' Make sure they want to do it
    sMsg = "Are you sure you wish to reset the password for user "
    sMsg = sMsg & msUserCode & "?"
    If MsgBox(sMsg, vbYesNo, "Reset MACRO Password") = vbNo Then Exit Sub

    ' Get login dates
    'Changed Mo 18/7/2001 field Password changed to UserPassword
    'Mo Morris 20/9/01 Db Audit (UserCode to UserName)
    sLoginDate = GetExpiredLogin
    sSQL = "UPDATE MacroUser SET " _
        & " UserPassword = 'macrotm', " _
        & " FirstLogin = " & sLoginDate & ", " _
        & " LastLogin = " & sLoginDate _
        & " WHERE UserName = '" & msUserCode & "'"
        
    SecurityADODBConnection.Execute sSQL, , adCmdText
    ' NCJ 17/7/00 - Added gLog message
    Call gLog("Reset Password", "The password for user " & msUserCode & " was reset.")

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdResetPwd_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub

'------------------------------------------------------------------------'
Private Sub Form_Load()
'------------------------------------------------------------------------'

'------------------------------------------------------------------------'

    Me.Icon = frmMenu.Icon
    FormCentre Me

End Sub

'--------------------------------------------------------------------------------
Private Sub lstUsers_Click()
'--------------------------------------------------------------------------------
' Click on user list - show user's full name
'--------------------------------------------------------------------------------
Dim nIndex As Integer

    On Error GoTo ErrHandler
    
    nIndex = lstUsers.ListIndex
    If nIndex > -1 Then     ' Was anything selected?
        ' Store as current user
        msUserCode = lstUsers.List(nIndex)
        txtUserName.Text = mcolUserNames(msUserCode)
        cmdResetPwd.Enabled = True
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "lstUsers_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub

'--------------------------------------------------------------------------------
Public Sub Display()
'--------------------------------------------------------------------------------
' Refresh ourselves and display
'--------------------------------------------------------------------------------

    'Reinitialise things
    lstUsers.Clear
    txtUserName.Text = ""
    cmdResetPwd.Enabled = False
    msUserCode = ""
    Set mcolUserNames = Nothing
    Set mcolUserNames = New Collection
    
    ' See if there are any users to reset
    If RefreshUsers Then
        FormCentre Me
        Me.Show vbModal
    Else
        MsgBox "There are no other users in the database.", vbOKOnly, "Reset Password"
    End If
    
End Sub

'--------------------------------------------------------------------------------
Private Function RefreshUsers() As Boolean
'--------------------------------------------------------------------------------
' Fill up the Users list box
' Return TRUE if there is at least one entry in the Users list,
' otherwise return FALSE
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim rsUserCodes As ADODB.Recordset
Dim sCode As String

    On Error GoTo ErrHandler
    
    ' Initialise to false
    RefreshUsers = False
    
    'Mo Morris 20/9/01 Db Audit (UserCode to UserName, UserName to UserNameFull)
    sSQL = "SELECT UserName, UserNameFull FROM MacroUser"
    Set rsUserCodes = New ADODB.Recordset
    rsUserCodes.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    While Not rsUserCodes.EOF
        sCode = rsUserCodes!UserName
        ' Add to list if not the current user
        If LCase$(sCode) <> LCase$(goUser.UserName) Then
            lstUsers.AddItem sCode
            ' Add Name to our collection, indexed by Code
            mcolUserNames.Add RemoveNull(rsUserCodes!UserNameFull), sCode
        End If
        rsUserCodes.MoveNext
    Wend
    
    rsUserCodes.Close
    Set rsUserCodes = Nothing
    
    If lstUsers.ListCount > 0 Then
        RefreshUsers = True
        ' Select the first item
        lstUsers.ListIndex = 0
        ' NB This generates a lstUsers_Click event
    End If
    
Exit Function
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "RefreshUsers")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Function

'--------------------------------------------------------------------------------
Private Function GetExpiredLogin() As String
'--------------------------------------------------------------------------------
' Get a date for login that means password will have expired
' Return result as "standardised" string for use in SQL
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim rsExpiry As ADODB.Recordset
Dim dblExpiry As Double
Dim dblLogin As Double

    On Error GoTo ErrHandler
    
    sSQL = "SELECT ExpiryPeriod FROM MacroPassword"
    Set rsExpiry = New ADODB.Recordset
    rsExpiry.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    ' Get password expiry period in days
    ' Retrieve it as a double for later arithmetic
    dblExpiry = CDbl(rsExpiry!ExpiryPeriod)
    
    rsExpiry.Close
    Set rsExpiry = Nothing
    
    ' Set login to ExpiryPeriod + 1 days before now
    dblLogin = CDbl(Now) - (dblExpiry + 1)
    ' Convert to "standard" string ready for use in SQL
    GetExpiredLogin = ConvertLocalNumToStandard(CStr(dblLogin))
    
Exit Function
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "GetExpiredLogin")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Function

'--------------------------------------------------------------------------------
Private Sub lstUsers_DblClick()
'--------------------------------------------------------------------------------
' Double click on list mimics "Reset Password" button
'--------------------------------------------------------------------------------

    Call cmdResetPwd_Click
    
End Sub
