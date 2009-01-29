VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   3735
   ClientLeft      =   6030
   ClientTop       =   6015
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraUser 
      Height          =   1095
      Left            =   60
      TabIndex        =   6
      Top             =   0
      Width           =   3255
      Begin VB.TextBox txtUserName 
         Height          =   345
         Left            =   1125
         TabIndex        =   0
         Top             =   240
         Width           =   2025
      End
      Begin VB.TextBox txtPassword 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1125
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   630
         Width           =   2025
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&User name"
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&Password"
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   645
         Width           =   900
      End
   End
   Begin VB.Frame fraDatabase 
      Caption         =   "Please select a database"
      Height          =   2055
      Left            =   60
      TabIndex        =   5
      Top             =   1140
      Width           =   3255
      Begin MSComctlLib.ListView lvwDatabase 
         Height          =   1695
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   2990
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Database"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   780
      TabIndex        =   2
      Top             =   3300
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2100
      TabIndex        =   3
      Top             =   3300
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       frmLogin.frm
'   Author:     Andrew Newbigging, June 1997
'   Purpose:    Allows user to enter user name and password.  If conditional compilation
'   constant NTSecurity  = 1 then the user's NT login name is automatically used
'   as the MACRO user name.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
 '  1       Andrew Newbigging       19/11/97
'   2       Andrew Newbigging       27/11/97
'   3       Andrew Newbigging       02/12/97
'   4       Andrew Newbigging       17/07/98
'   5       Andrew Newbigging       8/10/98     SPR 427
'           Reference to separate ImedSecurity component removed
'   6       Andrew Newbigging       9/11/98
'           Call to gUser.Login modified to force check of user password
'           Also, allow user to select a database using the keyboard
'   7       Andrew Newbigging       1/3/99
'           Conditional compilation constant NTSecurity replaced by gnSecurityMode global variable
'   8       PN  20/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   9       PN  26/09/99    Amended cmdOK_Click() for mtm1.6 changes
'   10      WillC 20/10/99  Addded Errhandler to SelectDatabase
'   NCJ 10 Nov 99 - Added error handlers
'   NCJ 9 Dec 99 - Added new checks for user function access
''  WillC   11/12/99
'          Changed the Following where present from Integer to Long  ClinicalTrialId
'           CRFPageId,VisitId,CRFElementID
'   NCJ 14 Dec 99 - Delay checking module access until after database selected
'   TA 20/04/2000 - appearance standardised, removed cmdSelectDatabase and cmdCancelDatbase
'                    and put code buhind standard OK and Cancel buttons
'   TA 25/04/2000   subclassing removed
'   TA 08/05/2000   new error handling when connecting to the macro database
'   TA 18/08/2000 UserHasAccessToModule function moved to gUser to help WWW development
'------------------------------------------------------------------------------------'
Option Explicit
Option Base 0
Option Compare Binary

Public LoginSucceeded As Boolean

#If ASE = 1 Then
  Private AseResult As Long
  Private hAseReader As Long
#End If

' PN 20/09/99
' new property to determine if user details are to be verified only
' with no furhter action
Private mbCheckPasswordOnly As Boolean
Private Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long)

'------------------------------------------------------------------------------------'
Public Property Let CheckPasswordOnly(bCheckPasswordOnly As Boolean)
'------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler
    
    mbCheckPasswordOnly = bCheckPasswordOnly
    txtUserName.Enabled = Not mbCheckPasswordOnly
    lblLabels(0).Enabled = Not mbCheckPasswordOnly
    
    Exit Property
    
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "CheckPasswordOnly")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Property

'------------------------------------------------------------------------------------'
Private Sub cmdCancel_Click()
'------------------------------------------------------------------------------------'
     
       'set the global var to false
    'to denote a failed login
    Call TidyUpAndHideWindow(False)

End Sub

'------------------------------------------------------------------------------------'
Public Sub TidyUpAndHideWindow(bLoginSucceeded As Boolean)
'------------------------------------------------------------------------------------'
    ' PN 20/09/99
    ' this should be set to false since normal login
    ' shall be a full login
    mbCheckPasswordOnly = False

    LoginSucceeded = bLoginSucceeded
    Me.Hide

End Sub

'------------------------------------------------------------------------------------'
Private Sub cmdOK_Click()
'------------------------------------------------------------------------------------'
'------------------------------------------------------------------------------------'
Dim nButtons As Integer

    On Error GoTo ErrHandler
    
   If lvwDatabase.Visible Then
       'user is selecting a database
       Call SelectDatabase
    Else
        'user is entering a username and password
        nButtons = vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal
        If txtUserName.Text = vbNullString Then
            MsgBox "Please enter a user name", nButtons, gsDIALOG_TITLE
            Call SetTextFocus
            
        ElseIf txtPassword.Text = vbNullString Then
            MsgBox "Please enter a password", nButtons, gsDIALOG_TITLE
            txtPassword.SetFocus
            
        Else
            HourglassOn
            '   ATN 9/11/98
            '   Check password parameter added to force check of user password
        
            ' PN 20/09/99
            ' only do the full login if required so pass new param mbCheckPasswordOnly
        
            Select Case gUser.login(txtUserName.Text, txtPassword.Text, True, mbCheckPasswordOnly)
            '   SPR 427 ATN 8/10/98
            '   Reference to Separate ImedSecurity module removed.
            Case LoginResult.Success
                ' PN 20/09/99
                 ' only do the full login if required so pass new param mbCheckPasswordOnly
                 If mbCheckPasswordOnly Then
                     LoginSucceeded = True
                     Me.Hide
                     
                 Else
                    ' NCJ - Delay setting of LoginSucceeded until they've chosen a database
                     Call RefreshUserDatabase
                     If lvwDatabase.ListItems.Count > 1 Then
                         Call Resize(True)
                         ' Wait until they choose a database
                     Else
                        ' There's only one database
                         If gUser.UserHasAccessToModule Then
                             LoginSucceeded = True
                         Else
                             ' We must boot them out
                             TidyUpAndHideWindow (False)
                         End If
    
                     End If
                 
                 End If
                
            '   SPR 427 ATN 8/10/98
            '   Reference to Separate ImedSecurity module removed.
            Case LoginResult.AccountDisabled
                MsgBox "Your account has been disabled.", nButtons, gsDIALOG_TITLE
                
                '   ATN 1/3/99
                '   Replaced conditional compilation with global variable
                If gnSecurityMode = SecurityMode.NTSeparatePassword Then
                    txtPassword.SetFocus
                ElseIf gnSecurityMode = SecurityMode.UsernamePassword Then
                    Call SetTextFocus
                End If
        
            '   SPR 427 ATN 8/10/98
            '   Reference to Separate ImedSecurity module removed.
            Case LoginResult.Failed
                MsgBox "Login failed", nButtons, gsDIALOG_TITLE
                
                '   ATN 1/3/99
                '   Replaced conditional compilation with global variable
                If gnSecurityMode = SecurityMode.NTSeparatePassword Then
                    txtPassword.SetFocus
                ElseIf gnSecurityMode = SecurityMode.UsernamePassword Then
                    Call SetTextFocus
                End If
    
            End Select
            
            HourglassOff
            
        End If
        
        ' PN 20/09/99
        ' this should be set to false since normal login
        ' shall be a full login
        mbCheckPasswordOnly = False
    End If
    
    Exit Sub
    
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdOK_Click")
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
Private Sub SetTextFocus()
'------------------------------------------------------------------------------------'
'
'------------------------------------------------------------------------------------'
     
    On Error GoTo ErrHandler
    
    If txtUserName.Enabled Then
        txtUserName.SetFocus
    Else
        txtPassword.SetFocus
    End If
    
    Exit Sub
    
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "SetTextFocus")
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
Private Sub Form_Load()
'------------------------------------------------------------------------------------'
'
'------------------------------------------------------------------------------------'
Dim msUserName As String
Dim mnUserNameLength As Long
Dim mnReturnCode As Long
    
    On Error GoTo ErrHandler
    
    Me.Icon = frmMenu.Icon
    
    ' NCJ 9/12/99 - Default to false
    LoginSucceeded = False
    
    ' Turn on key preview for form, so that F1 (Help) can be trapped by form
    Me.KeyPreview = True
    Me.BackColor = glFormColour
    
    txtUserName.Text = vbNullString
    txtPassword.Text = vbNullString
    
    Call Resize(False)
    
    FormCentre Me
    
    '   ATN 1/3/99
    '   Replaced conditional compilation with global variable
    If gnSecurityMode = SecurityMode.NTSeparatePassword Then
        mnUserNameLength = 199
        msUserName = String(200, 0)
        mnReturnCode = GetUserName(msUserName, mnUserNameLength)
        txtUserName.Text = Left(msUserName, mnUserNameLength)
        txtUserName.Enabled = False
        
    End If
    
    #If ASE = 1 Then
        If GetASEUserId <> txtUserName.Text Then
            MsgBox "The card in the socket does not match your user name."
            ExitMACRO
        
        End If
    
    #End If
    
    ' PN 20/09/99
    ' this should be set to false since normal login
    ' shall be a full login
    mbCheckPasswordOnly = False

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


'------------------------------------------------------------------------------------'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'------------------------------------------------------------------------------------'
'
'------------------------------------------------------------------------------------'

    If KeyCode = vbKeyF1 Then               ' Show user guide
        'ShowDocument Me.hWnd, gsMACROUserGuidePath
        
        'REM 07/12/01 - New Call for MACRO Help
        Call MACROHelp(Me.hWnd, App.Title)
    End If

End Sub

'------------------------------------------------------------------------------------'
Private Sub lvwDatabase_DblClick()
'------------------------------------------------------------------------------------'
 '
'------------------------------------------------------------------------------------'

    Call SelectDatabase
    
End Sub


'------------------------------------------------------------------------------------'
Private Sub lvwDatabase_KeyDown(KeyCode As Integer, Shift As Integer)
'------------------------------------------------------------------------------------'
'   ATN 9/11/98
'   Check for Cr added to allow selection of database using keyboard
'------------------------------------------------------------------------------------'
    
    If KeyCode = Asc(vbCr) Then
        Call SelectDatabase
    End If

End Sub

'------------------------------------------------------------------------------------'
Private Sub txtPassword_GotFocus()
'------------------------------------------------------------------------------------'
'
'------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler
    
    SendKeys "{Home}+{End}"
       
    Exit Sub
    
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtPassword_GotFocus")
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
Private Sub txtUserName_GotFocus()
'------------------------------------------------------------------------------------'
 
'------------------------------------------------------------------------------------'
    
    On Error GoTo ErrHandler
    
    SendKeys "{Home}+{End}"
       
    Exit Sub
    
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtUserName_GotFocus")
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
Private Sub RefreshUserDatabase()
'------------------------------------------------------------------------------------'
'------------------------------------------------------------------------------------'
Dim msUserDatabase As Variant

    On Error GoTo ErrHandler
    
    lvwDatabase.ListItems.Clear
    
    For Each msUserDatabase In gUser.UserDatabases
    
        lvwDatabase.ListItems.Add , msUserDatabase, msUserDatabase
    
    Next
    
    If lvwDatabase.ListItems.Count = 1 Then
        gUser.SetDatabasePath lvwDatabase.ListItems(1)
        Me.Hide
    ElseIf lvwDatabase.ListItems.Count = 0 Then
        MsgBox "You do not have permission to access any MACRO databases", vbOKOnly + vbApplicationModal
        cmdOK.Enabled = False
    Else
        lvwDatabase.ListItems(1).Selected = True
        cmdOK.Enabled = True
    End If
       
    Exit Sub
    
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "RefreshUserDatabase")
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
Private Sub SelectDatabase()
'----------------------------------------------------------------------------------------'
' User is selecting a database
' WillC added handler for when a user chooses a Security database
' instead of a Macro database
' NCJ 14 Dec 99 - Check that user has access to current module on this database
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler


    'TA SR3428: do not use infermed error handling for SetDatabasePath
    On Error GoTo ErrHandler2
    gUser.SetDatabasePath lvwDatabase.SelectedItem
    ' NCJ - We can't check user functions until the database has been set up,
    ' so do it now
    
    On Error GoTo ErrHandler
    If gUser.UserHasAccessToModule Then
        LoginSucceeded = True
    End If
    
    Me.Hide

    Exit Sub
    
ErrHandler:
    Select Case Err.Number
        Case 400
            Exit Sub
        Case Else
            Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "SelectDatabase")
                Case OnErrorAction.Ignore
                    Resume Next
                Case OnErrorAction.Retry
                    Resume
                Case OnErrorAction.QuitMACRO
                    Call ExitMACRO
                    Call MACROEnd
            End Select
    End Select
Exit Sub

ErrHandler2:
    'TA SR3428: new error handling
    Select Case Err.Number
        Case 400  ' cannot display form, it's already shown modally
            Exit Sub
        Case Else
            Call MsgBox("The following problem occured when attempting to connect to the MACRO database:" & vbCrLf & vbCrLf & Err.Description, _
                vbOKOnly + vbCritical, "Cannot connect to MACRO database")
    End Select
Exit Sub

End Sub

'----------------------------------------------------------------------------------------'
Private Sub Resize(fShowDatabases As Boolean)
'----------------------------------------------------------------------------------------'
' TA 20/04/2000 - resize form and enable controls according to whether
'                   the user is entering password or choosing a database
'----------------------------------------------------------------------------------------'
    
    fraDatabase.Visible = fShowDatabases
    fraUser.Enabled = Not fShowDatabases
    
    If fShowDatabases Then
        Me.Height = 4140
        cmdCancel.Top = fraDatabase.Top + fraDatabase.Height + 120
        lvwDatabase.SetFocus
    Else
        Me.Height = 2040
        cmdCancel.Top = fraUser.Top + fraUser.Height + 120
    End If
    
    cmdOK.Top = cmdCancel.Top
    
End Sub




#If ASE = 1 Then
Private Function GetASEUserId() As String

 Const WRITE_READ_DATA_SIZE = 8
 Const KEY_SIZE = 8
 Const FILE_ID = 6
 Const FILE_SIZE = 8
 Const OFFSET = 0
 
Dim dwActiveProtocol As Long
Dim hAseCard As Long
ReDim chWriteBuffer(WRITE_READ_DATA_SIZE) As Byte
ReDim chreadbuffer(WRITE_READ_DATA_SIZE) As Byte
Dim MainKey(KEY_SIZE) As Byte
Dim chWriteKey(KEY_SIZE) As Byte
Dim WriteBufferTmp As String
Dim ReadBufferTmp As String
Dim wActualDataRead As Integer
Dim IO As ASEIO_T0
Dim i As Integer
Dim CardCaps As HLCARDCAPS
Dim FileProperties  As FileProperties
Dim no As Variant

Dim msASEUserId As String

     On Error GoTo ErrHandler
    
   '-----------------------------------------------------------------------------------------
    ' Initialize Main key to HFF's
    '-----------------------------------------------------------------------------------------
    For i = 0 To KEY_SIZE
        MainKey(i) = &HFF
    Next i
    
    '-----------------------------------------------------------------------------------------
    ' Initializing file properties (ID: 6, 8 bytes, no acces rules)
    '-----------------------------------------------------------------------------------------
    FileProperties.wID = FILE_ID
    FileProperties.wBytesAllocated = FILE_SIZE
    FileProperties.wWriteConditions = AC_NONE
    FileProperties.wReadConditions = AC_NONE
    
    
    '-----------------------------------------------------------------------------------------
    '   Open the first reader that is in the ASE database
    '-----------------------------------------------------------------------------------------
    'Debug.Print "Open the ASEDrive."
    DoEvents
    AseResult = ASEReaderOpenByNameNull(0, hAseReader)
    If (ReportResult(AseResult) = 0) Then
        End
    End If
    
    
    '-----------------------------------------------------------------------------------------
    '   Power the ISO7816 T=0 card
    '-----------------------------------------------------------------------------------------
    'Debug.Print "Power the ISO7816 T=0 card."
    DoEvents
    
    AseResult = ASECardPowerOn( _
                                    hAseReader, _
                                    MAIN_SOCKET, _
                                    CARD_POWER_UP, _
                                    PROTOCOL_CPU7816_T0, _
                                    dwActiveProtocol, _
                                    hAseCard)
    
    Do Until ReportResult(AseResult) <> 0
        AseResult = ASECardPowerOn( _
                                    hAseReader, _
                                    MAIN_SOCKET, _
                                    CARD_POWER_UP, _
                                    PROTOCOL_CPU7816_T0, _
                                    dwActiveProtocol, _
                                    hAseCard)
    Loop
        
    
    '-----------------------------------------------------------------------------------------
    '   Check if the card in the socket is a T=0 card
    '-----------------------------------------------------------------------------------------
    If (dwActiveProtocol <> PROTOCOL_CPU7816_T0) Then
        MsgBox ("Error: The card in the socket is not a T=0 card !!!")
        Call CloseReader
        End
    End If


    '-----------------------------------------------------------------------------------------
    '   Selecting Card Level
    '-----------------------------------------------------------------------------------------
    'Debug.Print "Selecting Card Level"
    DoEvents
    
    AseResult = ASEHLSelectCardLevel(hAseCard)
    If (ReportResult(AseResult) = 0) Then
       Call CloseReader
       End
    End If
    

    '-----------------------------------------------------------------------------------------
    '   Get card capabilities in order to retrieve information for further use.
    '-----------------------------------------------------------------------------------------
    'Debug.Print "Get card capabilities "
    DoEvents
    AseResult = ASEHLGetCardCaps(hAseCard, CardCaps)
    If (ReportResult(AseResult) = 0) Then
       Call CloseReader
       End
    End If
    
    '-----------------------------------------------------------------------------------------
    '   Creation of file 6
    '-----------------------------------------------------------------------------------------
    
'    Debug.Print "Creation of file 6"
'    DoEvents
'    AseResult = ASEHLCreateFile(hAseCard, MainKey(0), FileProperties)
'    If (ReportResult(AseResult) = 0) Then
'       Call CloseReader
'       End
'    End If
   
    '-----------------------------------------------------------------------------------------
    '   Open file number 6
    '-----------------------------------------------------------------------------------------
    'Debug.Print "Open file number 6."
    DoEvents
    AseResult = ASEHLOpenFile(hAseCard, FILE_ID)
    If (ReportResult(AseResult) = 0) Then
        Call CloseReader
        End
    End If
    
    '-----------------------------------------------------------------------------------------
    '   Write to file number 6
    '-----------------------------------------------------------------------------------------
'    Debug.Print "Write to file number 6: 01234567"
'    DoEvents
    
    '-----------------------------------------------------------------------------------------
    ' Set the buffer
    '-----------------------------------------------------------------------------------------
'    For I = 0 To WRITE_READ_DATA_SIZE - 1
'        chWriteBuffer(I) = I
'    Next I
'        chWriteBuffer(0) = Asc("a")
'        chWriteBuffer(1) = Asc("n")
'        chWriteBuffer(2) = Asc("d")
'        chWriteBuffer(3) = Asc("r")
'        chWriteBuffer(4) = Asc("e")
'        chWriteBuffer(5) = Asc("w")
'        chWriteBuffer(6) = Asc("n")
'        chWriteBuffer(7) = Asc(" ")
'
'    If (CardCaps.dwSecuredWriting = 1) Then
'        For I = 0 To KEY_SIZE
'            chWriteKey(I) = &HFF
'        Next I
''        For I = 0 To WRITE_READ_DATA_SIZE - 1
'            AseResult = ASEHLWrite(hAseCard, WRITE_READ_DATA_SIZE, OFFSET, chWriteBuffer(0), chWriteKey(0))
'            If (ReportResult(AseResult) = 0) Then
'                Call CloseReader
'                End
'            End If
''        Next
'    Else
'        no = Null
'        AseResult = ASEHLWriteUnprotect(hAseCard, WRITE_READ_DATA_SIZE, OFFSET, chWriteBuffer(0), 0)
'        If (ReportResult(AseResult) = 0) Then
'            Call CloseReader
'            End
'        End If
'    End If
'
    
    '-----------------------------------------------------------------------------------------
    '   Initialize the read buffer - set it to 0
    '-----------------------------------------------------------------------------------------
    For i = 0 To WRITE_READ_DATA_SIZE - 1
        chreadbuffer(i) = 0
    Next i

    msASEUserId = ""

    '-----------------------------------------------------------------------------------------
    '   Read from file number 6
    '-----------------------------------------------------------------------------------------
    Debug.Print "Read from file number 6."
    DoEvents
    
'    For I = 0 To WRITE_READ_DATA_SIZE - 1
        AseResult = ASEHLRead(hAseCard, WRITE_READ_DATA_SIZE, OFFSET, chreadbuffer(0))
        If (ReportResult(AseResult) = 0) Then
            Call CloseReader
            End
        End If
 '   Next

    For i = 0 To WRITE_READ_DATA_SIZE - 1
        msASEUserId = msASEUserId & Chr(chreadbuffer(i))
    Next

    GetASEUserId = RTrim(msASEUserId)
    
    '-----------------------------------------------------------------------------------------
    '   Check if chWriteBuffer and chReadBuffer are identical
    '-----------------------------------------------------------------------------------------
'    Debug.Print "Check that the buffer are identical."
'    DoEvents
'    WriteBufferTmp = chWriteBuffer
'    ReadBufferTmp = chReadBuffer
'    If (WriteBufferTmp <> ReadBufferTmp) Then
'        MsgBox "Error: The buffers are not identical."
'        Call CloseReader
'        End
'    End If
'    Debug.Print "The buffers are identicale."
    
    Call CloseReader
       
    Exit Function
    
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "GetASEUserID")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select

End Function

'=============================================================================================
' @NAME: ReportResult
' @DESC: This fuction displays the error that has occured
' @ARGS: AseResult As Long - result of previous call to API
' @RTRN: True or False( depends on wether the result is positive or not )
'=============================================================================================
Private Function ReportResult(AseResult As Long) As Boolean

Dim bResult As Boolean

    ' Let errors be handed by calling routine - NCJ 10/11/99
    
    Select Case AseResult
        Case ASEERR_SUCCESS
            Debug.Print (" The function succeeded")
            
        Case ASEERR_FAIL
            MsgBox ("The function failed")
    
        Case ASEERR_READER_ALREADY_OPEN
            MsgBox ("The specified reader is already opened")
    
        Case ASEERR_TIMEOUT
            MsgBox ("The function returned after timeout")
    
        Case ASEERR_WRONG_READER_NAME
            MsgBox ("No reader with the specified name exists")
    
        Case ASEERR_READER_OPEN_ERROR
            MsgBox ("The specified reader could not be opened")
    
        Case ASEERR_READER_COMM_ERROR
            MsgBox ("Reader communication error")
    
        Case ASEERR_MAX_READERS_ALREADY_OPEN
            MsgBox ("The maximum number of readers is already opened")
    
        Case ASEERR_INVALID_READER_HANDLE
            MsgBox ("The specified reader handle is invalid")
    
        Case ASEERR_SYSTEM_ERROR
            MsgBox ("General system error has occurred")
        
        Case ASEERR_INVALID_SOCKET
            MsgBox ("The specified socket is invalid")
    
        Case ASEERR_OPERATION_TIMEOUT
            MsgBox ("Blocking operation has been canceled after timeout")
    
        Case ASEERR_OPERATION_CANCELED
            MsgBox ("Blocking operation has been canceled by the user")
    
        Case ASEERR_INVALID_PARAMETERS
            MsgBox ("One or more of the specified parameters is invalid")
    
        Case ASEERR_PROTOCOL_NOT_SUPPORTED
            MsgBox ("The specified protocol is not supported")
    
        Case ASEERR_CARD_COMM_ERROR
            MsgBox ("Card communication error")
    
        Case ASEERR_CARD_NOT_PRESENT
            MsgBox ("Please insert your SmartCard into the card reader.")
    
        Case ASEERR_CARD_NOT_POWERED
            MsgBox ("The card is not powered on")
    
        Case ASEERR_IFSD_OVERFLOW
            MsgBox ("The command's data length is too big")
    
        Case ASEERR_CARD_INVALID_PARAMETER
            MsgBox ("One or more of the card parameters is invalid")
    
        Case ASEERR_INVALID_CARD_HANDLE
            MsgBox ("The specified card handle is invalid")
    
        Case ASEERR_NOT_INSTALLED
            MsgBox ("Ase is not installed in your system")
    
        Case ASEERR_COMMAND_NOT_SUPPORTED
            MsgBox ("This command is not supported. Read manual for details")
    
        Case ASEERR_MEMORY_CARD_ERROR
            MsgBox ("A memory card error has occurred. Read manual for details")
    
        Case ASEERR_NO_RTC
            MsgBox ("There is no RTC on this reader")
    
        Case ASEERR_WRONG_ACTIVE_PROTOCOL
            MsgBox ("This command can not work with the current protocol")
    
        Case ASEERR_NO_READER_AT_PORT
            MsgBox ("There is no ASE reader in the specified port")
    
        Case ASEERR_CARD_ALREADY_POWERED
            MsgBox ("Card is already powered")
    
        Case ASEERR_NO_HL_CARD_SUPPORT
            MsgBox ("The card has no ASE high level API support")
    
        Case ASEERR_CANT_LOAD_CARD_DLL
            MsgBox ("Can not load the high level API of the current card")
    
        Case ASEERR_WRONG_PASSWORD
            MsgBox ("Wrong password")
    
        Case ASEWRN_SERIAL_NUMBER_MISMATCH
            MsgBox ("The reader serial number does not match the registered one")
            
        'High Level errors
        Case ASEHLERR_UNSUPPORTED_CARD
            MsgBox ("ASEHLERR_UNSUPPORTED_CARD")
            
        Case ASEHLERR_KEY_EXISTS
            MsgBox ("ASEHLERR_KEY_EXISTS")
            
        Case ASEHLERR_INVALID_ID
            MsgBox ("ASEHLERR_INVALID_ID")
            
        Case ASEHLERR_INVALID_OFFSET
            MsgBox ("ASEHLERR_INVALID_OFFSET")

        Case ASEHLERR_UNFULFILLED_CONDITIONS
            MsgBox ("ASEHLERR_UNFULFILLED_CONDITIONS")

        Case ASEHLERR_INVALID_LENGTH
            MsgBox ("ASEHLERR_INVALID_LENGTH")

        Case ASEHLERR_WRONG_KEY
            MsgBox ("ASEHLERR_WRONG_KEY")

        Case ASEHLERR_BLOCKED
            MsgBox ("ASEHLERR_BLOCKED")
            
        Case ASEHLERR_SECURE_WRITE_UNSUPPORTED
            MsgBox ("ASEHLERR_SECURE_WRITE_UNSUPPORTED")

        Case ASEHLERR_CARD_MEMORY_PROBLEM
            MsgBox ("ASEHLERR_CARD_MEMORY_PROBLEM")

            
        Case ASEHLERR_INVALID_KEYREF
            MsgBox ("ASEHLERR_INVALID_KEYREF")

        Case ASEHLERR_UNSUPPORTED_FUNCTION
            MsgBox ("ASEHLERR_UNSUPPORTED_FUNCTION")

        Case ASEHLERR_KEY_NOT_EXIST
            MsgBox ("ASEHLERR_KEY_NOT_EXIST")

        Case ASEHLERR_CARD_INSUFFICIENT_MEMORY
            MsgBox ("ASEHLERR_CARD_INSUFFICIENT_MEMORY")
                                                  
        Case ASEHLERR_ID_ALREADY_EXISTS
            MsgBox ("ASEHLERR_ID_ALREADY_EXISTS")
                                                  
        Case ASEHLERR_API_FATAL_ERROR
            MsgBox ("ASEHLERR_API_FATAL_ERROR")
            
        Case ASEHLERR_API
            MsgBox ("ASEHLERR_API")
            
        Case ASEHLERR_INCORRECT_PARAMETER
            MsgBox ("ASEHLERR_INCORRECT_PARAMETER")
            
        Case ASEHLERR_INVALID_FILE
            MsgBox ("ASEHLERR_INVALID_FILE")
            
        Case ASEHLERR_FILE_NOT_OPEN
            MsgBox ("ASEHLERR_FILE_NOT_OPEN")
            
        Case ASEHLERR_NO_MORE_CHANGES
            MsgBox ("ASEHLERR_NO_MORE_CHANGES")
            
        Case ASEHLERR_FAILURE
            MsgBox ("ASEHLERR_FAILURE")
            
        Case ASEHLERR_CARD_FATAL_ERROR
            MsgBox ("ASEHLERR_CARD_FATAL_ERROR")
            
        Case ASEHLERR_CARD_ERROR
            MsgBox ("ASEHLERR_CARD_ERROR")
            
        Case Else
            MsgBox ("Unknown ASE error " & AseResult)
    
    End Select

    If AseResult = ASEERR_SUCCESS Or AseResult = ASEERR_CARD_ALREADY_POWERED Then
        bResult = True
    Else
        bResult = False
    End If


    ReportResult = bResult

End Function

'=============================================================================================
' @NAME: CloseReader
' @DESC: Closes reader handler
' @ARGS: NONE
' @RTRN: NONE
'=============================================================================================
Private Sub CloseReader()
    Debug.Print "Close the reader."
    AseResult = ASEReaderClose(hAseReader)
    ReportResult (AseResult)
End Sub
#End If


