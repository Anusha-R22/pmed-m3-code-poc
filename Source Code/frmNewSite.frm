VERSION 5.00
Begin VB.Form frmNewSite 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Site"
   ClientHeight    =   2550
   ClientLeft      =   7845
   ClientTop       =   7110
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optRemote 
      Caption         =   "Remote"
      Height          =   315
      Left            =   2520
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin VB.OptionButton optServer 
      Caption         =   "Server/Web"
      Height          =   315
      Left            =   1140
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   7
      Top             =   60
      Width           =   5955
      Begin VB.ComboBox cboCountry 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1560
         Width           =   4875
      End
      Begin VB.TextBox txtCode 
         Height          =   315
         Left            =   1020
         MaxLength       =   8
         TabIndex        =   0
         ToolTipText     =   "Site code may be up to 8 characters long and may not contain spaces"
         Top             =   240
         Width           =   4875
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1020
         MaxLength       =   255
         TabIndex        =   1
         ToolTipText     =   "Site description may be up to 255 characters"
         Top             =   630
         Width           =   4875
      End
      Begin VB.Label lblCountry 
         Alignment       =   1  'Right Justify
         Caption         =   "Country"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1620
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Location"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1140
         Width           =   795
      End
      Begin VB.Label lblCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Code"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   270
         Width           =   435
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   "Description"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   690
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4860
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "frmNewSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1998-2004. All Rights Reserved
'   File:       frmNewSite.frm
'   Author:     Paul Norris, 21/07/1999
'   Purpose:    User can add sites with a code and a site name
'
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
'   Revisions:
'   1   PN  10/09/99    Updated code to conform to VB standards doc version 1.0
'   PN  21/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   WillC  10/11/99  Added the error handlers
'   NCJ 18 Jan 00   SR2206 Centre form on frmSiteAdmin form
'   NCJ 29 Apr 00   SRs 3289, 3357 (site codes/descriptions)
'                   Removed subclassing; added correct HourGlass calls
'   NCJ 27 Sep 00   Tidied up tooltip text
'   ZA 18/09/2002   Update List_sites.js if user creates a new site
'   ZA 27/09/2002   Added System Locales, Countries and language list
'   ASH 13/1/2003  Added optional parameter to GetMacroDBSetting in Form_Load
'   NCJ 15 Nov 04 - (Isse 2433) In cmdOK_Click, ignore roles with no function codes
'--------------------------------------------------------------------------------

Option Explicit

Private WithEvents moSite As clsSite
Attribute moSite.VB_VarHelpID = -1
Private mbIsNewSite As Boolean
Private mbIsLoading As Boolean
' NCJ 29/4/00 - Store if user changed site description
Private mbUserChangedName As Boolean
' NCJ 15 Nov 04 - Moved some module variables to Form_Load (because only used there)
'ASH 11/12/2002
Private msDatabase As String
Private mconMACRO As ADODB.Connection

'--------------------------------------------------------------------------------
Public Property Let SiteCode(sSiteCode As String)
'--------------------------------------------------------------------------------
    
    moSite.Code = sSiteCode

End Property

'--------------------------------------------------------------------------------
Public Property Get IsNewSite() As Boolean
'--------------------------------------------------------------------------------
    
    IsNewSite = mbIsNewSite

End Property

'--------------------------------------------------------------------------------
Public Property Let IsNewSite(bIsNewSite As Boolean)
'--------------------------------------------------------------------------------
    
    If moSite Is Nothing Then
        Set moSite = New clsSite
    End If

    mbIsNewSite = bIsNewSite

End Property

'--------------------------------------------------------------------------------
Private Sub EnableOK(bIsValid As Boolean)
'--------------------------------------------------------------------------------
' handle the enabling of the ok button
'--------------------------------------------------------------------------------
     
     On Error GoTo ErrHandler
 
    ' the Code and the Name are both mandatory
    If bIsValid Then
        ' both are populated
        cmdOK.Enabled = True
    Else
        ' one or other is not populated
        cmdOK.Enabled = False
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

'--------------------------------------------------------------------------------
Private Sub cboCountry_Click()
'--------------------------------------------------------------------------------
'set the country code
'--------------------------------------------------------------------------------
    If cboCountry.ListIndex > -1 Then
        moSite.CountryCode = cboCountry.ItemData(cboCountry.ListIndex)
    Else
        EnableOK False
    End If
End Sub

'--------------------------------------------------------------------------------
Private Sub cboSiteLocale_Click()
'--------------------------------------------------------------------------------
'set the site locale
'--------------------------------------------------------------------------------

'    If cboSiteLocale.ListIndex > -1 Then
'        moSite.SiteLocale = cboSiteLocale.ItemData(cboSiteLocale.ListIndex)
'    Else
'        EnableOK False
'    End If
    
End Sub

'--------------------------------------------------------------------------------
Private Sub cboTimeZone_Click()
'--------------------------------------------------------------------------------
'set the time zone
'--------------------------------------------------------------------------------

'    If cboTimeZone.ListIndex > -1 Then
'        moSite.TimeZone = cboTimeZone.ItemData(cboTimeZone.ListIndex)
'    Else
'        EnableOK False
'    End If
End Sub

'--------------------------------------------------------------------------------
Private Sub cmdCancel_Click()
'--------------------------------------------------------------------------------
' cancel the action (either edit site or add site)
'--------------------------------------------------------------------------------
     On Error GoTo ErrHandler
 
    ' cancel the action and exit
    Set moSite = Nothing
    
    Unload Me
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "cmdCancel_Click")
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
Private Sub cmdOK_Click()
'--------------------------------------------------------------------------------
' accept the action (either edit site or add site)
' NCJ 15 Nov 04 - Ignore roles with no function codes (Isse 2433)
'--------------------------------------------------------------------------------
Dim sMSG As String
Dim oSystemMessage As SysMessages
Dim sMessageParameters As String
Dim vRoleCodes As Variant
Dim sRoleCode As String
Dim sRoleDescription As String
Dim nEnabled As Integer
Dim vRole As Variant
Dim vFunctionCodes As Variant
Dim sFunctionCodes As String
Dim sFunctionCode As String
Dim i As Integer
Dim j As Integer
Dim vUsers As Variant
Dim sUsername As String
Dim vUserDetails As Variant
Dim sSiteCode As String
Dim sStudyCode As String
Dim nTypeOfInstal As Integer
Dim vUserRoles As Variant
Dim nSysAdmin As Integer
    
    On Error GoTo ErrHandler
    
    Call HourglassOn
    
    ' check if the site can be saved
    If moSite.Save = ESiteAlreadyExists Then
        'TA 20/03/2000   SR3264 ammended to show site code the same
        sMSG = "A site with the code '" & moSite.Code & "' already exists. Please try a different code"
        MsgBox sMSG, vbOKOnly + vbExclamation, gsDIALOG_TITLE
            
    Else
        'ZA 18/09/2002 - update the script file
        CreateSitesList
             
        'Distribute Password Policy
        Set oSystemMessage = New SysMessages
        sMessageParameters = goUser.PasswordPolicy.MinPswdLength & gsPARAMSEPARATOR _
                           & goUser.PasswordPolicy.MaxPswdLength & gsPARAMSEPARATOR _
                           & goUser.PasswordPolicy.ExpiryPeriod & gsPARAMSEPARATOR _
                           & -CInt(goUser.PasswordPolicy.EnforceMixedCase) & gsPARAMSEPARATOR _
                           & -CInt(goUser.PasswordPolicy.EnforceDigit) & gsPARAMSEPARATOR _
                           & -CInt(goUser.PasswordPolicy.AllowRepeatChars) & gsPARAMSEPARATOR _
                           & -CInt(goUser.PasswordPolicy.AllowUserName) & gsPARAMSEPARATOR _
                           & goUser.PasswordPolicy.PasswordHistory & gsPARAMSEPARATOR _
                           & goUser.PasswordPolicy.PasswordRetries
        Call oSystemMessage.AddNewSystemMessage(MacroADODBConnection, ExchangeMessageType.PasswordPolicy, goUser.UserName, goUser.UserName, "Password Policy", sMessageParameters, moSite.Code)
        
        'Distribute  Roles
        Set oSystemMessage = New SysMessages
        'returns all the RoleCodes in the security database
        vRoleCodes = RoleCodes
        If Not IsNull(vRoleCodes) Then
            For i = 0 To UBound(vRoleCodes, 2)
                sRoleCode = vRoleCodes(0, i)
                sRoleDescription = vRoleCodes(1, i)
                nEnabled = vRoleCodes(2, i)
                nSysAdmin = vRoleCodes(3, i)
                'return all the function codes for a specific role
                vFunctionCodes = FunctionCodes(sRoleCode)
                ' NCJ 15 Nov 04 - Ignore roles with no function codes (Isse 2433)
                If Not IsNull(vFunctionCodes) Then
                    sFunctionCodes = ""
                    For j = 0 To UBound(vFunctionCodes, 2)
                        sFunctionCode = vFunctionCodes(0, j)
                        sFunctionCodes = sFunctionCodes & sFunctionCode
                        If j <> UBound(vFunctionCodes, 2) Then
                            sFunctionCodes = sFunctionCodes & gsSEPARATOR
                        End If
                    Next
                    sMessageParameters = sRoleCode & gsPARAMSEPARATOR & sRoleDescription & gsPARAMSEPARATOR _
                                & nEnabled & gsPARAMSEPARATOR & nSysAdmin & gsPARAMSEPARATOR & sFunctionCodes
                    Call oSystemMessage.AddNewSystemMessage(MacroADODBConnection, ExchangeMessageType.Role, _
                                    goUser.UserName, goUser.UserName, "Role Management", sMessageParameters, moSite.Code)
                End If
            Next
        End If
        
        'Distribute Users
        Set oSystemMessage = New SysMessages
        vUsers = UserWithAllSitesUserRoles
        If Not IsNull(vUsers) Then
            For i = 0 To UBound(vUsers, 2)
                sUsername = vUsers(0, i)
                If Not ExcludeUserRDE(sUsername) Then 'don't distribute rde
                    vUserDetails = UserDetails(sUsername)
                    If Not IsNull(vUserDetails) Then
                        'REM 24/03/04 - elements 4,5 & 8 are doubles from the DB
                        sMessageParameters = vUserDetails(0, 0) & gsPARAMSEPARATOR _
                                           & vUserDetails(1, 0) & gsPARAMSEPARATOR _
                                           & vUserDetails(2, 0) & gsPARAMSEPARATOR _
                                           & vUserDetails(3, 0) & gsPARAMSEPARATOR _
                                           & LocalNumToStandard(vUserDetails(4, 0)) & gsPARAMSEPARATOR _
                                           & LocalNumToStandard(vUserDetails(5, 0)) & gsPARAMSEPARATOR _
                                           & vUserDetails(7, 0) & gsPARAMSEPARATOR _
                                           & LocalNumToStandard(vUserDetails(8, 0)) & gsPARAMSEPARATOR _
                                           & vUserDetails(9, 0) & gsPARAMSEPARATOR _
                                           & eUserDetails.udEditUser
                        Call oSystemMessage.AddNewSystemMessage(MacroADODBConnection, ExchangeMessageType.User, goUser.UserName, goUser.UserName, "User Details", sMessageParameters, moSite.Code)
                    End If
                End If
            Next
        End If
        
        'Distribute UserRoles
        Set oSystemMessage = New SysMessages
        'if there are users with all sites permission then distribute them
        If Not IsNull(vUsers) Then
            For i = 0 To UBound(vUsers, 2)
                sUsername = vUsers(0, i)
                If Not ExcludeUserRDE(sUsername) Then 'don't distribute rde user roles
                    vUserRoles = UserRoles(sUsername)
                    If Not IsNull(vUserRoles) Then
                        For j = 0 To UBound(vUserRoles, 2)
                            sUsername = vUserRoles(0, j)
                            sRoleCode = vUserRoles(1, j)
                            sStudyCode = vUserRoles(2, j)
                            sSiteCode = vUserRoles(3, j)
                            nTypeOfInstal = vUserRoles(4, j)
                        
                            sMessageParameters = sUsername & gsPARAMSEPARATOR _
                                           & sRoleCode & gsPARAMSEPARATOR _
                                           & sStudyCode & gsPARAMSEPARATOR _
                                           & sSiteCode & gsPARAMSEPARATOR _
                                           & nTypeOfInstal & gsPARAMSEPARATOR _
                                           & eUserRole.urAdd
                            Call oSystemMessage.AddNewSystemMessage(MacroADODBConnection, ExchangeMessageType.UserRole, goUser.UserName, goUser.UserName, "New User Role", sMessageParameters, moSite.Code)
                        
                        Next
                    End If
                End If
            Next
        End If
        
        Set oSystemMessage = Nothing
        
        ' exit this window
        Set moSite = Nothing
        Unload Me
        
    End If
    

    
    Call HourglassOff
    
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

'--------------------------------------------------------------------------------
Private Function UserRoles(sUsername As String) As Variant
'--------------------------------------------------------------------------------
'REM 22/11/02
'Returns all the UserRoles for a specific user
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim rsUserRoles As ADODB.Recordset

    On Error GoTo ErrLabel

    sSQL = "SELECT * FROM UserRole" _
        & " WHERE UserName = '" & sUsername & "'"
    Set rsUserRoles = New ADODB.Recordset
    rsUserRoles.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
        
    If rsUserRoles.RecordCount > 0 Then
        UserRoles = rsUserRoles.GetRows
    Else
        UserRoles = Null
    End If
    
    rsUserRoles.Clone
    Set rsUserRoles = Nothing

Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "frmNewSite.UserRoles"
End Function

'--------------------------------------------------------------------------------
Private Function UserDetails(sUsername As String) As Variant
'--------------------------------------------------------------------------------
'REM 22/11/02
'Returns a users details from the MACRO User table
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim rsUser As ADODB.Recordset

    On Error GoTo ErrLabel

    sSQL = "SELECT * FROM MACROUser" _
        & " WHERE UserName = '" & sUsername & "'"
    Set rsUser = New ADODB.Recordset
    rsUser.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText

    If rsUser.RecordCount > 0 Then
        UserDetails = rsUser.GetRows
    Else
        UserDetails = Null
    End If

    rsUser.Close
    Set rsUser = Nothing

Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "frmNewSite.UserDetails"
End Function

'--------------------------------------------------------------------------------
Private Function UserWithAllSitesUserRoles() As Variant
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim rsUsers As ADODB.Recordset

    On Error GoTo ErrLabel
    
    sSQL = "SELECT DISTINCT UserName FROM UserRole" _
        & " WHERE SiteCode = 'AllSites'"
    Set rsUsers = New ADODB.Recordset
    rsUsers.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsUsers.RecordCount > 0 Then
        UserWithAllSitesUserRoles = rsUsers.GetRows
    Else
        UserWithAllSitesUserRoles = Null
    End If

Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "frmNewSite.UserWithAllSitesUserRoles"
End Function

'--------------------------------------------------------------------------------
Private Function RoleCodes() As Variant
'--------------------------------------------------------------------------------
'REM 22/11/02
'Returns all the RoleCodes from the Role table in the security database
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim rsRoleCodes As ADODB.Recordset

    On Error GoTo ErrLabel

    sSQL = "SELECT * FROM Role"
    Set rsRoleCodes = New ADODB.Recordset
    rsRoleCodes.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText

    If rsRoleCodes.RecordCount > 0 Then
        RoleCodes = rsRoleCodes.GetRows
    Else
        RoleCodes = Null
    End If
    
    rsRoleCodes.Close
    Set rsRoleCodes = Nothing

Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "frmNewSite.RoleCodes"
End Function

'--------------------------------------------------------------------------------
Private Function FunctionCodes(sRoleCode As String) As Variant
'--------------------------------------------------------------------------------
'REM 22/11/02
'Returns all the function codes for a sepcific Role
'--------------------------------------------------------------------------------
Dim sSQL As String
Dim rsFunctionCodes As ADODB.Recordset

    On Error GoTo ErrLabel
    
    sSQL = "SELECT FunctionCode" _
        & " FROM RoleFunction" _
        & " WHERE RoleCode = '" & sRoleCode & "'"
    Set rsFunctionCodes = New ADODB.Recordset
    rsFunctionCodes.Open sSQL, SecurityADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsFunctionCodes.RecordCount > 0 Then
        FunctionCodes = rsFunctionCodes.GetRows
    Else
        FunctionCodes = Null
    End If
    
    rsFunctionCodes.Close
    Set rsFunctionCodes = Nothing

Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "frmNewSite.FunctionCodes"
End Function

'--------------------------------------------------------------------------------
Private Sub Form_Load()
'--------------------------------------------------------------------------------
' when the form is loaded the IsNewSite property is called and the object
' is instantiated here it is loaded with the sitecode if it exists
' and does nothing if it is new
'--------------------------------------------------------------------------------
' REVISIONS
' DPH 06/08/2002 - Added Server / Remote Location
' NCJ 15 Nov 04 - Moved module variables to here and made them local
'--------------------------------------------------------------------------------
Dim bServerLocation As Boolean
Dim n As Integer
Dim vDicKeys As Variant
Dim sMessage As String
Dim bLoad As Boolean
Dim oDatabase As MACROUserBS30.Database
Dim sConnectionString As String

    On Error GoTo ErrHandler
    
    Me.Icon = frmMenu.Icon
     
    mbUserChangedName = False
    
    mbIsLoading = True
    
    Set oDatabase = New MACROUserBS30.Database
    bLoad = oDatabase.Load(SecurityADODBConnection, goUser.UserName, msDatabase, "", False, sMessage)
    sConnectionString = oDatabase.ConnectionString
    Set mconMACRO = New ADODB.Connection
    mconMACRO.Open sConnectionString
    mconMACRO.CursorLocation = adUseClient
    
    ' NCJ 15 Nov 04 - Tidy up
    Set oDatabase = Nothing
    
    LoadCountries
    
    If Not mbIsNewSite And moSite.Code <> vbNullString Then
        ' editing an existing site -  read the site properties
        moSite.Load
        txtCode.Text = moSite.Code
        txtName.Text = moSite.Name
        Me.Caption = "Edit Site - [" & moSite.Code & "]"
        txtCode.Enabled = False
        Select Case moSite.Location
            Case SiteLocation.ESiteServer
                optServer.Value = True
            Case SiteLocation.ESiteRemote
                optRemote.Value = True
            Case SiteLocation.ESiteNoLocation
                optServer.Value = False
                optRemote.Value = False
        End Select
        
        If moSite.CountryCode = NULL_LONG Then
            cboCountry.ListIndex = -1
        Else
            SetSiteCountry moSite.CountryCode
        End If
        
    Else
        Me.Caption = "Create Site " & "[" & goUser.DatabaseCode & "]"

    End If
    
    ' DPH 07/08/2002 - Check if a Server or remote site installation
    If GetMacroDBSetting("datatransfer", "dbtype", mconMACRO, gsSERVER) = "site" Then
        Call EnableLocation(False)
    Else
        Call EnableLocation(True)
    End If
    
    ' NCJ 18/1/00 SR 2206 Centre ourselves on the frmSiteAdmin window
    Me.Left = FrmSiteAdmin.Left + (FrmSiteAdmin.Width - Me.Width) \ 2
    Me.Top = FrmSiteAdmin.Top + (FrmSiteAdmin.Height - Me.Height) \ 2
    
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
'--------------------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'--------------------------------------------------------------------------------
'clears out the dictionary now
'--------------------------------------------------------------------------------
    
    gdicSystemLocales.RemoveAll
    
End Sub

'--------------------------------------------------------------------------------
Private Sub moSite_IsValid(bIsValid As Boolean)
'--------------------------------------------------------------------------------
' the site has raised its isvalid event so respond by enabling/disabling
' the ok button
'--------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Call EnableOK(bIsValid)
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "moSite_IsValid")
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
Private Sub optRemote_Click()
'--------------------------------------------------------------------------------
' Selection of the Remote Site Location Option
'--------------------------------------------------------------------------------
    
    If Not mbIsLoading Then
        moSite.Location = SiteLocation.ESiteRemote
    End If
    
End Sub

'--------------------------------------------------------------------------------
Private Sub optServer_Click()
'--------------------------------------------------------------------------------
' Selection of the Server Site Location Option
'--------------------------------------------------------------------------------

    If Not mbIsLoading Then
        moSite.Location = SiteLocation.ESiteServer
    End If
    
End Sub

'--------------------------------------------------------------------------------
Private Sub txtCode_Change()
'--------------------------------------------------------------------------------
' the clsSite object encapsulates the rules for a site code
' capture the error when the input is invalid
' NCJ 29/4/00 - SRs 3289 and 3357
' REM 23/01/03 - added LCase to Site Code text box so site code is always lower case
'--------------------------------------------------------------------------------
 Dim iPos As Integer
    
    On Error GoTo InvalidChar
    
    iPos = txtCode.SelStart
    
    If Not mbIsLoading Then
        txtCode.Text = LCase(txtCode.Text)
        moSite.Code = txtCode.Text
        
        ' populate the Site Name field with the Site Code if the name is blank
        ' or if the user hasn't changed the name
        On Error GoTo FieldEmpty
        
        If txtName.Text = vbNullString Or Not mbUserChangedName Then
            txtName.Text = txtCode.Text
        End If
        
    End If
    
    'Exit Sub
    
InvalidChar:
    txtCode = moSite.Code
    If iPos > 0 Then
        txtCode.SelStart = iPos
    End If
        
FieldEmpty:

End Sub

'--------------------------------------------------------------------------------
Private Sub txtName_Change()
'--------------------------------------------------------------------------------
' the clsSite object encapsulates the rules for a site name
' capture the error when the input is invalid
'--------------------------------------------------------------------------------
Dim iPos As Integer
    
    On Error GoTo InvalidChar
    
    iPos = txtName.SelStart
    If Not mbIsLoading Then
        moSite.Name = txtName.Text
        
    End If
    
    Exit Sub
    
InvalidChar:
    txtName = moSite.Name
    If iPos > 0 Then
        txtName.SelStart = iPos - 1
    End If

End Sub

'--------------------------------------------------------------------------------
Private Sub txtName_KeyPress(KeyAscii As Integer)
'--------------------------------------------------------------------------------
' The user typed in this field so switch off automatic tracking of txtCode field
'--------------------------------------------------------------------------------

    mbUserChangedName = True

End Sub

'--------------------------------------------------------------------------------
Private Sub EnableLocation(bEnable As Boolean)
'--------------------------------------------------------------------------------
' Enable Location for Server Installations
'--------------------------------------------------------------------------------

    optServer.Enabled = bEnable
    optRemote.Enabled = bEnable
    
End Sub

'--------------------------------------------------------------------------------
Private Sub LoadCountries()
'--------------------------------------------------------------------------------
'load country name from MACROCountry table
'--------------------------------------------------------------------------------
Dim oRS As ADODB.Recordset
Dim sQuery As String

    On Error GoTo ErrLabel
    
    Set oRS = New ADODB.Recordset
    sQuery = "select CountryId, CountryDescription from MACROCountry order by CountryDescription ASC"
    
    oRS.Open sQuery, mconMACRO, adOpenKeyset, adLockOptimistic
    
    If Not oRS.EOF Then
        Do While Not oRS.EOF
            cboCountry.AddItem oRS.Fields("CountryDescription").Value
            cboCountry.ItemData(cboCountry.NewIndex) = oRS.Fields("CountryId").Value
            oRS.MoveNext
        Loop
    End If
    
    Set oRS = Nothing
    
    Exit Sub
    
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "LoadCountries", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------------------
Private Sub LoadTimeZones()
'--------------------------------------------------------------------------------
'load country name from MACROCountry table
'--------------------------------------------------------------------------------
'Dim oRS As ADODB.Recordset
'Dim sQuery As String
'
'    On Error GoTo errlabel
'
'    Set oRS = New ADODB.Recordset
'    sQuery = "select TZId, Description from MACROTimeZone order by OffsetMins asc"
'
'    oRS.Open sQuery, mconMACRO, adOpenKeyset, adLockOptimistic
'
'    If Not oRS.EOF Then
'        Do While Not oRS.EOF
'            cboTimeZone.AddItem oRS.Fields("Description").Value
'            cboTimeZone.ItemData(cboTimeZone.NewIndex) = oRS.Fields("TZId").Value
'            oRS.MoveNext
'        Loop
'    End If
'
'    Set oRS = Nothing
'
'    Exit Sub
    
ErrLabel:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "LoadTimeZones", Err.Source) = Retry Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------------------
Private Sub SetSiteLocale(ByVal lSiteLocale As Long)
'--------------------------------------------------------------------------------
'determines the site locale from the combo that was selected for the site
'--------------------------------------------------------------------------------
'Dim n As Integer
'
'    For n = 0 To cboSiteLocale.ListCount - 1
'        If cboSiteLocale.ItemData(n) = lSiteLocale Then
'            cboSiteLocale.ListIndex = n
'            Exit Sub
'        End If
'    Next n
End Sub

'--------------------------------------------------------------------------------
Private Sub SetSiteCountry(ByVal lSiteCountry As Long)
'--------------------------------------------------------------------------------
'determines the site country from the combo that was selected for the site
'--------------------------------------------------------------------------------
Dim n As Integer

    For n = 0 To cboCountry.ListCount - 1
        If cboCountry.ItemData(n) = lSiteCountry Then
            cboCountry.ListIndex = n
            Exit Sub
        End If
    Next n
End Sub

'--------------------------------------------------------------------------------
Private Sub SetSiteTimeZone(ByVal lSiteTimeZone As Long)
'--------------------------------------------------------------------------------
'determines the site country from the combo that was selected for the site
'--------------------------------------------------------------------------------
'Dim n As Integer
'
'    For n = 0 To cboTimeZone.ListCount - 1
'        If cboTimeZone.ItemData(n) = lSiteTimeZone Then
'            cboTimeZone.ListIndex = n
'            Exit Sub
'        End If
'    Next n
End Sub

'---------------------------------------------------------------------
Public Property Get Database() As String
'---------------------------------------------------------------------
' NCJ 15 Nov 04 - Corrected this (but don't think it's ever used...)
'---------------------------------------------------------------------
    
    Database = msDatabase
'    msDatabase = Database

End Property

'---------------------------------------------------------------------
Public Property Let Database(sDatabase As String)
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------
    
     msDatabase = sDatabase

End Property
