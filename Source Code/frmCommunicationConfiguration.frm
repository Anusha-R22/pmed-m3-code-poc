VERSION 5.00
Object = "{0E56FD71-943D-11D2-BA66-0040053687FE}#1.0#0"; "DartWeb.dll"
Begin VB.Form frmCommunicationConfiguration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Communication Configuration"
   ClientHeight    =   6720
   ClientLeft      =   2370
   ClientTop       =   1830
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   4245
   Begin VB.OptionButton optNoTransfer 
      Caption         =   "No data transfer"
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   5460
      Width           =   1455
   End
   Begin VB.OptionButton optRequestTransfer 
      Caption         =   "Transfer on request"
      Height          =   315
      Left            =   1560
      TabIndex        =   12
      Top             =   5820
      Width           =   1935
   End
   Begin VB.CommandButton cmdTestCon 
      Caption         =   "&Test"
      Height          =   345
      Left            =   1560
      TabIndex        =   10
      Top             =   5040
      Width           =   1125
   End
   Begin VB.Frame Frame4 
      Caption         =   "Effective Dates"
      Height          =   1035
      Left            =   60
      TabIndex        =   27
      Top             =   3900
      Width           =   4095
      Begin VB.TextBox txtEffectiveFrom 
         BackColor       =   &H0000FFFF&
         Height          =   285
         Left            =   1500
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtEffectiveTo 
         BackColor       =   &H0000FFFF&
         Height          =   285
         Left            =   1500
         TabIndex        =   9
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   280
         Width           =   1005
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   650
         Width           =   1005
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Advanced Settings (HTTPS)"
      Height          =   1395
      Left            =   60
      TabIndex        =   23
      Top             =   2460
      Width           =   4100
      Begin VB.TextBox txtPortNumber 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Tag             =   "ValidateZero"
         Top             =   300
         Width           =   735
      End
      Begin VB.TextBox txtHTTP 
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Tag             =   "ValidateZero"
         Top             =   660
         Width           =   2295
      End
      Begin VB.TextBox txtProxyServer 
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Tag             =   "ValidateZero"
         Top             =   1020
         Width           =   2295
      End
      Begin VB.Label lblPortNumber 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Port Number"
         Height          =   255
         Left            =   0
         TabIndex        =   26
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label lblHTTP 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "HTTP Address"
         Height          =   255
         Left            =   0
         TabIndex        =   25
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label lblProxy 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proxy Server"
         Height          =   255
         Left            =   0
         TabIndex        =   24
         Top             =   1020
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Basic Settings"
      Height          =   1335
      Left            =   60
      TabIndex        =   19
      Top             =   1080
      Width           =   4100
      Begin VB.TextBox txtUserName 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Tag             =   "ValidateZero"
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   3
         Tag             =   "ValidateZero"
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtConfirmPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   4
         Tag             =   "ValidateZero"
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lblUserName 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblPassword 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   255
         Left            =   0
         TabIndex        =   21
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Site Details"
      Height          =   975
      Left            =   60
      TabIndex        =   16
      Top             =   60
      Width           =   4100
      Begin VB.TextBox txtSite 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtTrialOffice 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label lblSite 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Site"
         Enabled         =   0   'False
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblTrialOffice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Study Office"
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   345
      Left            =   2820
      TabIndex        =   15
      Top             =   6300
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   1560
      TabIndex        =   14
      Top             =   6300
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   345
      Left            =   300
      TabIndex        =   13
      Top             =   6300
      Width           =   1125
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Test Connection"
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   5085
      Width           =   1200
   End
   Begin DartWebCtl.Http Http1 
      Left            =   360
      OleObjectBlob   =   "frmCommunicationConfiguration.frx":0000
      Top             =   5520
   End
End
Attribute VB_Name = "frmCommunicationConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       frmCommunicationConfiguration.frm
'   Author:     Paul Norris 22/07/99
'   Purpose:    Configuration window for the communication settings
'               stored in the TrialOffice table in Macro.mdb
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   PN  20/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   PN  30/09/99    Changed class names
'                   clsCommunicationData to clsCommunication
'                   because prog id is too long with original name
'   ATN 11/12/99    Replaced form icon and added field for ProxyServer setting, which was missing
'  WillC 11 / 12 / 99
'          Changed the Following where present from Integer to Long  ClinicalTrialId
'           CRFPageId,VisitId,CRFElementID
'   Mo Morris 20/12/99
'           edits to txtEffectivFrom/To now enable the Ok and Apply buttons
'   NCJ 21/12/99 Tidied up date field validation and button enabling
'   ATN 12/1/2000   In ValidateandSave routine, refresh the clsCommunication attached to frmMenu
'   NCJ 20/1/00     SR2672 Changed "Trial office" to "Study Office"
'   NCJ 18/2/00 - Date fields are now automatically formatted (see clsCommunication)
'   NCJ 7/3/00 - SRs2780,3110 Major rewriting using new clsComms class
'   TA 25/04/2000   subclassing removed
'   WillC 3/8/00    Changed Errhandlers
'   TA 26/9/01: Changes following db audit
'   DPH 9/1/2002 Proxy Server setting not being shown
'   DPH 23/07/2002  - CBBL 2.2.19.23 Mark site as now being remote
'   REM 28/05/03 - Added a IsNew property to the form so can tell if creating a new setting or editing old one
'                - Reloaded comms objects when user clicks apply for a new comm setting
'------------------------------------------------------------------------------------

Option Explicit

' The Comm object for this form
Private WithEvents moCommunicationConfig As clsCommunication
Attribute moCommunicationConfig.VB_VarHelpID = -1
' The collection of Comm objects
Private mcolCommConfigs As clsComms

Private mbIsLoading As Boolean

' NCJ 21/12/99 - Background colour for invalid fields
Private Const mlInvalidColour = &HFFFF&

Private Const msTEST_CONNECTION_URL = "Test_Connection.asp"

'REM 28/05/03 - is new property
Private mbIsNew As Boolean

'---------------------------------------------------------------------
Public Property Get IsNew() As Boolean
'---------------------------------------------------------------------
    
    IsNew = mbIsNew

End Property

'---------------------------------------------------------------------
Public Property Let IsNew(bIsNew As Boolean)
'---------------------------------------------------------------------

    mbIsNew = bIsNew

End Property

'---------------------------------------------------------------------
Public Property Get Component() As clsCommunication
'---------------------------------------------------------------------

    Set Component = moCommunicationConfig

End Property

'------------------------------------------------------------------------------'
Public Property Set Component(oConfigSetting As clsCommunication)
'---------------------------------------------------------------------
' Set the communication object for the form
' Refresh fields with relevant data
'---------------------------------------------------------------------

    Set moCommunicationConfig = oConfigSetting
    Call SetNewCommunicationConfig

End Property

'---------------------------------------------------------------------
Public Property Set CommConfigs(oCommConfigs As clsComms)
'---------------------------------------------------------------------
' Set the collection of communication objects for the form
'---------------------------------------------------------------------

    Set mcolCommConfigs = oCommConfigs

End Property

'---------------------------------------------------------------------
Public Property Get CommConfigs() As clsComms
'---------------------------------------------------------------------
' Get the collection of communication objects for the form
'---------------------------------------------------------------------

    Set CommConfigs = mcolCommConfigs

End Property

'------------------------------------------------------------------------------'
Private Sub cmdTestCon_Click()
'------------------------------------------------------------------------------'
'REM 26/11/02
'Test the connection to the server using Test asp page
'------------------------------------------------------------------------------'
Dim sTestMessage As String
Dim sHTTPAddress As String
Dim vTestMsg As Variant
Dim sMessage As String
Dim bError As Boolean

    On Error GoTo ErrConnect
    
    bError = False
    
    sHTTPAddress = txtHTTP.Text
    
    sTestMessage = TestDataHTTP(sHTTPAddress & msTEST_CONNECTION_URL, 60)
    
    vTestMsg = Split(sTestMessage, gsMSGSEPARATOR)
    
    If vTestMsg(0) = "Success" Then
        'connection the the server succeeded
        sMessage = "Connection to MACRO web server successful!" & vbCrLf
        'was error in creating the connection object on the server
        If vTestMsg(1) <> "" Then
            sMessage = sMessage & vTestMsg(1) & vbCrLf
            bError = True
        End If
        'was an error in connection to the server database
        If vTestMsg(2) <> "" Then
            sMessage = sMessage & vTestMsg(2)
            bError = True
        End If
        
        If bError Then
            Call DialogError(sMessage)
        Else
            Call DialogInformation("Connection to MACRO server database successful!")
        End If
        
    Else
        Call DialogError("Connection to the MACRO web server failed.  Please contact your system administrator.")
    End If
    
Exit Sub
ErrConnect:
    Call DialogError("Connection failed")
End Sub

'---------------------------------------------------------------------
Private Function TestDataHTTP(ByVal sAddress As String, _
                                   ByVal vTimeout As Integer) As String
'---------------------------------------------------------------------
' Post Data using Dart Control catch error but raise
'---------------------------------------------------------------------
Dim sData As String
Dim sURLParamData As String
Dim sUser As String
Dim sPassword As String
Dim nPortNumber As Integer

On Error GoTo ErrHandler
    
    sData = ""
    sURLParamData = ""
    sUser = txtUserName.Text
    sPassword = txtPassword.Text
    nPortNumber = txtPortNumber.Text
    
    ' Timeout
    Call SetTimeoutHTTP(vTimeout)
    
    ' Set URL (with port)
    Http1.URL = frmDataTransfer.AddPortToHTTPString(sAddress, nPortNumber)
    
    Http1.Post sURLParamData, , sData, , sUser, sPassword
    
    ' Return collected string
    TestDataHTTP = sData
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|PostDataHTTP"
End Function

'---------------------------------------------------------------------
Private Sub SetTimeoutHTTP(ByVal nTimeout As Integer)
'---------------------------------------------------------------------
' Set Timeout on HTTP control
'---------------------------------------------------------------------

    ' Timeout in seconds (needs to be milliseconds)
    Http1.Timeout = CLng(nTimeout) * 1000
    
End Sub

'------------------------------------------------------------------------------'
Private Sub Form_Unload(Cancel As Integer)
'------------------------------------------------------------------------------'

    Set moCommunicationConfig = Nothing
    Set mcolCommConfigs = Nothing

End Sub

'------------------------------------------------------------------------------'
Private Function ValidateData() As String
'------------------------------------------------------------------------------'
' Validation of data entered
' NCJ 28/2/00 - This routine now returns relevant error message
' Returns empty string for no errors
'------------------------------------------------------------------------------'
Dim sMsg As String

    On Error GoTo ErrHandler

    sMsg = ""
    
    ' Check passwords
    If txtConfirmPassword.Text <> txtPassword.Text Then
        sMsg = "The passwords entered are not the same."
    Else
        ' Check date fields
        sMsg = mcolCommConfigs.ValidateCommRecord(moCommunicationConfig)
    End If
    
    ValidateData = sMsg
    
Exit Function
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "ValidateData")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Function

'------------------------------------------------------------------------------'
Private Sub EnableOK(bEnable As Boolean)
'------------------------------------------------------------------------------'

    On Error GoTo ErrHandler

    If bEnable Then
        ' all populated
        cmdOK.Enabled = True
        cmdApply.Enabled = True
        cmdTestCon.Enabled = True
    Else
        ' one or other is not populated
        cmdOK.Enabled = False
        cmdApply.Enabled = False
        cmdTestCon.Enabled = False
    End If
    
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "EnableOK")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'------------------------------------------------------------------------------'
Private Sub ValidateAndSave(Optional bUnload As Boolean = False)
'------------------------------------------------------------------------------'
' REVISIONS
' DPH 23/07/2002  - CBBL 2.2.19.23 Mark site as now being remote
'------------------------------------------------------------------------------'
Dim sMsg As String

    On Error GoTo ErrHandler

    ' NCJ 28/2/00 - ValidateData now returns an error message (if any)
    sMsg = ValidateData

    If sMsg = "" Then
        ' Warn if already expired
        If moCommunicationConfig.DblEffectiveTo < CDbl(CLng(CDbl(Now))) Then
            sMsg = "The Effective To date of " & vbNewLine
            sMsg = sMsg & "    " & moCommunicationConfig.EffectiveTo & vbNewLine
            sMsg = sMsg & "is before today's date, meaning this setting has already expired." & vbNewLine
            sMsg = sMsg & "You will not be able to make any more changes to this record after saving it." & vbNewLine
            sMsg = sMsg & "Are you sure you wish to continue?"
            If MsgBox(sMsg, vbOKCancel) = vbCancel Then
                Exit Sub
            End If
        End If
        ' All valid so apply edits
        moCommunicationConfig.Save
        '   ATN 12/1/2000
        '   Update the frmMenu communication class
        ' NCJ 7/3/00 - New routine
        frmMenu.SetUpTrialOffice
        
        ' DPH 23/07/2002  - CBBL 2.2.19.23 Mark site as now being remote
        gblnRemoteSite = True
        
        Call EnableWriteOnceFields(False)
        
        If bUnload Then
            Unload Me
        End If
    Else
        ' Validation failed
        MsgBox sMsg, vbCritical + vbOKOnly, gsDIALOG_TITLE
        Call EnableOK(False)
    End If

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "ValidateAndSave")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'------------------------------------------------------------------------------'
Private Sub cmdApply_Click()
'------------------------------------------------------------------------------'
Dim sKey As String

    Call ValidateAndSave
    
    'REM 28/05/03 - If creating a new setting thenreload objects in case user then edits and clicks apply again
    If mbIsNew Then
        Set mcolCommConfigs = New clsComms
        mcolCommConfigs.Load
        Set moCommunicationConfig = New clsCommunication
        sKey = txtTrialOffice.Text & CDbl(CDate(txtEffectiveFrom.Text)) & CDbl(CDate(txtEffectiveTo.Text))
        Set moCommunicationConfig = mcolCommConfigs.Item(sKey)
        mbIsNew = False
    End If
    
End Sub

'------------------------------------------------------------------------------'
Private Sub cmdCancel_Click()
'------------------------------------------------------------------------------'

    Unload Me
    
End Sub

'------------------------------------------------------------------------------'
Private Sub cmdOK_Click()
'------------------------------------------------------------------------------'

    Call ValidateAndSave(True)
    
End Sub

'------------------------------------------------------------------------------'
Private Sub SetNewCommunicationConfig()
'------------------------------------------------------------------------------'
' Set a new communication object - fill in form fields
' REVISIONS
' DPH 9/1/2002 - Corrected Proxy server setting not being displayed
'------------------------------------------------------------------------------'

    mbIsLoading = True
    
    With moCommunicationConfig
        txtUserName = .User
        txtPassword = .Password
        txtConfirmPassword = .Password
        txtHTTP = .HTTPAddress
        txtPortNumber = .PortNumber
        ' DPH 9/1/2002 - Added missing Proxy Server setting
        txtProxyServer.Text = .ProxyServer
        txtSite = .Site
        txtTrialOffice = .TrialOffice
        ' NCJ 18/2/00 - EffectiveTo and EffectiveFrom fields are now automatically formatted correctly
        txtEffectiveFrom.Text = .EffectiveFrom
        txtEffectiveTo.Text = .EffectiveTo
'   ATN 11/12/99
'   Replaced check box with option buttons
        Select Case .TransferData
        Case 0
            optNoTransfer.Value = True
        Case 1
            optRequestTransfer.Value = True
        Case Else
            optNoTransfer.Value = True
        End Select
'        chkTransferData.Value = IIf(.TransferData = vbNullString, 0, .TransferData)
'        Call EnableOK(.IsValid)
        ' No changes to begin with
        Call EnableOK(False)
    End With
    
    If txtSite.Text = "" Then
        txtSite.Text = SetSiteFromMACRODBSettings
    End If
    
    Call EnableWriteOnceFields(txtSite.Text = vbNullString)
    
    mbIsLoading = False

End Sub

'------------------------------------------------------------------------------'
Private Function SetSiteFromMACRODBSettings() As String
'------------------------------------------------------------------------------'
'REM 26/11/02
'Returns the site code using the MACRODBSettings table if there is one set
'------------------------------------------------------------------------------'
Dim sSQL As String
Dim rsSiteCode As ADODB.Recordset
    
    On Error GoTo ErrHandler

    sSQL = "SELECT SettingValue FROM MACRODBSetting" _
        & " WHERE SettingSection = 'datatransfer'" _
        & " AND SettingKey = 'dbsitename'"
    Set rsSiteCode = New ADODB.Recordset
    rsSiteCode.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If rsSiteCode.RecordCount > 0 Then
        SetSiteFromMACRODBSettings = rsSiteCode!SettingValue
    Else
        SetSiteFromMACRODBSettings = ""
    End If
    
    rsSiteCode.Close
    Set rsSiteCode = Nothing
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "frmCommunicationConfiguration.SetSiteFromMACRODBSettings"
End Function

'------------------------------------------------------------------------------'
Private Sub Form_Load()
'------------------------------------------------------------------------------'
' Once-only initialisations
'------------------------------------------------------------------------------'
Dim sToolTip As String

   On Error GoTo ErrHandler

    Me.Icon = frmMenu.Icon
    
    mbIsLoading = True
    Me.BackColor = glFormColour
    
    sToolTip = "Please enter the date in the format " & frmMenu.DefaultDateFormat
    txtEffectiveFrom.TooltipText = sToolTip
    txtEffectiveTo.TooltipText = sToolTip
    
    FormCentre Me
    
    mbIsLoading = False
    
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

'------------------------------------------------------------------------------'
Private Sub EnableWriteOnceFields(bEnable As Boolean)
'------------------------------------------------------------------------------'
    
    On Error GoTo ErrHandler

    lblSite.Enabled = bEnable
    txtSite.Enabled = bEnable

    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "EnableWriteOnceFields")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'------------------------------------------------------------------------------'
Private Sub moCommunicationConfig_IsValid(bIsValid As Boolean)
'------------------------------------------------------------------------------'
    
    On Error GoTo ErrHandler

   If moCommunicationConfig.HasChanges Then
        Call EnableOK(bIsValid)
    Else
        Call EnableOK(False)
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "moCommunicationConfig_IsValid")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'------------------------------------------------------------------------------'
Private Sub optNoTransfer_Click()
'------------------------------------------------------------------------------'
    
    On Error GoTo ErrHandler

    If Not mbIsLoading Then
        moCommunicationConfig.TransferData = 0
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "optNoTransfer_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

'------------------------------------------------------------------------------'
Private Sub optRequestTransfer_Click()
'------------------------------------------------------------------------------'

    On Error GoTo ErrHandler

    If Not mbIsLoading Then
        moCommunicationConfig.TransferData = 1
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "optRequestTransfer_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub

'------------------------------------------------------------------------------'
Private Sub txtConfirmPassword_Change()
'------------------------------------------------------------------------------'

    Call EnableOK(moCommunicationConfig.IsValid)
    
End Sub

'------------------------------------------------------------------------------'
Private Sub txtEffectiveFrom_Change()
'------------------------------------------------------------------------------'

    ' NCJ 21/12/99 - Add validation of value
    On Error GoTo InvalidDate
    If txtEffectiveFrom.Text > "" Then
        ' If date is not valid, this assignment raises an error
        moCommunicationConfig.EffectiveFrom = txtEffectiveFrom.Text
        'txtEffectiveFrom.BackColor = vbWhite
        
        'ASH 17/1/2003 set backcolour to system colour
        txtEffectiveFrom.BackColor = vbWindowBackground
        
        ' Enable buttons if everything else is OK
        Call EnableOK(moCommunicationConfig.IsValid)
    Else
        txtEffectiveFrom.BackColor = mlInvalidColour
        Call EnableOK(False)
    End If
    
    Exit Sub
    
InvalidDate:
    txtEffectiveFrom.BackColor = mlInvalidColour
    Call EnableOK(False)
    
End Sub

'------------------------------------------------------------------------------'
Private Sub txtEffectiveFrom_LostFocus()
'------------------------------------------------------------------------------'
' NCJ 21/12/99 - Validation checks are now done in the Change event
'------------------------------------------------------------------------------'
   
End Sub

'------------------------------------------------------------------------------'
Private Sub txtEffectiveTo_Change()
'------------------------------------------------------------------------------'

    ' NCJ 21/12/99 - Add validation of value
    On Error GoTo InvalidDate
    If txtEffectiveTo.Text > "" Then
        ' If date is not valid, this assignment raises an error
        moCommunicationConfig.EffectiveTo = txtEffectiveTo.Text
        'txtEffectiveTo.BackColor = vbWhite
        
        'ASH 17/1/2003 set backcolour to system colour
        txtEffectiveTo.BackColor = vbWindowBackground

        Call EnableOK(moCommunicationConfig.IsValid)
    Else
        txtEffectiveTo.BackColor = mlInvalidColour
        Call EnableOK(False)
    End If
    Exit Sub
    
InvalidDate:
    txtEffectiveTo.BackColor = mlInvalidColour
    Call EnableOK(False)

End Sub

'------------------------------------------------------------------------------'
Private Sub txtEffectiveTo_LostFocus()
'------------------------------------------------------------------------------'
' NCJ 21/12/99 - Validation checks are now done in the Change event
'------------------------------------------------------------------------------'

End Sub

'------------------------------------------------------------------------------'
Private Sub txtHTTP_Change()
'------------------------------------------------------------------------------'

    If Not mbIsLoading Then
        Call TextChange(txtHTTP, moCommunicationConfig, "HTTPAddress")
    End If
    
End Sub

'------------------------------------------------------------------------------'
Private Sub txtHTTP_LostFocus()
'------------------------------------------------------------------------------'

    txtHTTP = TextLostFocus(txtHTTP, moCommunicationConfig, "HTTPAddress")
    
End Sub

'------------------------------------------------------------------------------'
Private Sub txtProxyServer_Change()
'------------------------------------------------------------------------------'

    If Not mbIsLoading Then
        Call TextChange(txtProxyServer, moCommunicationConfig, "ProxyServer")
    End If
    
End Sub

'------------------------------------------------------------------------------'
Private Sub txtProxyServer_LostFocus()
'------------------------------------------------------------------------------'

    txtProxyServer = TextLostFocus(txtProxyServer, moCommunicationConfig, "ProxyServer")
    
End Sub

'------------------------------------------------------------------------------'
Private Sub txtPassword_Change()
'------------------------------------------------------------------------------'

    If Not mbIsLoading Then
        Call TextChange(txtPassword, moCommunicationConfig, "Password")
    End If
    
End Sub

'------------------------------------------------------------------------------'
Private Sub txtPassword_LostFocus()
'------------------------------------------------------------------------------'

    txtPassword = TextLostFocus(txtPassword, moCommunicationConfig, "Password")
    
End Sub

'------------------------------------------------------------------------------'
Private Sub txtPortNumber_Change()
'------------------------------------------------------------------------------'

    If Not mbIsLoading Then
        Call TextChange(txtPortNumber, moCommunicationConfig, "PortNumber")
    End If

End Sub

'------------------------------------------------------------------------------'
Private Sub txtPortNumber_LostFocus()
'------------------------------------------------------------------------------'

    txtPortNumber = TextLostFocus(txtPortNumber, moCommunicationConfig, "PortNumber")
    
End Sub

'------------------------------------------------------------------------------'
Private Sub txtSite_Change()
'------------------------------------------------------------------------------'

    If Not mbIsLoading Then
        Call TextChange(txtSite, moCommunicationConfig, "Site")
    End If
  
    
    
End Sub

'------------------------------------------------------------------------------'
Private Sub txtSite_LostFocus()
'------------------------------------------------------------------------------'
    
    On Error GoTo ErrHandler

    txtSite = TextLostFocus(txtSite, moCommunicationConfig, "Site")
        
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtSite_LostFocus")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   

End Sub


'------------------------------------------------------------------------------'
Private Sub txtTrialOffice_Change()
'------------------------------------------------------------------------------'

    If Not mbIsLoading Then
        Call TextChange(txtTrialOffice, moCommunicationConfig, "TrialOffice")
    End If
    
End Sub

'------------------------------------------------------------------------------'
Private Sub txtTrialOffice_LostFocus()
'------------------------------------------------------------------------------'

    On Error GoTo ErrHandler

    txtTrialOffice = TextLostFocus(txtTrialOffice, moCommunicationConfig, "TrialOffice")
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtTrialOffice_LostFocus")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'------------------------------------------------------------------------------'
Private Sub txtUserName_Change()
'------------------------------------------------------------------------------'
    
    If Not mbIsLoading Then
        Call TextChange(txtUserName, moCommunicationConfig, "User")
    End If
    
End Sub

'------------------------------------------------------------------------------'
Private Sub txtUserName_LostFocus()
'------------------------------------------------------------------------------'

    On Error GoTo ErrHandler

    txtUserName = TextLostFocus(txtUserName, moCommunicationConfig, "User")
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtUserName_LostFocus")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

