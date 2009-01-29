VERSION 5.00
Begin VB.Form frmRRConnection 
   Caption         =   "Registration/Randomisation Server Settings"
   ClientHeight    =   3465
   ClientLeft      =   6330
   ClientTop       =   5955
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   7485
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6195
      TabIndex        =   11
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4860
      TabIndex        =   10
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Frame fraConnection 
      Caption         =   "Remote Server Connection Details"
      Height          =   1515
      Left            =   60
      TabIndex        =   12
      Top             =   1380
      Width           =   7335
      Begin VB.TextBox txtPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   4920
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   660
         Width           =   2295
      End
      Begin VB.TextBox txtProxyServer 
         Height          =   315
         Left            =   1320
         MaxLength       =   255
         TabIndex        =   9
         Top             =   1080
         Width           =   5895
      End
      Begin VB.TextBox txtHTTPAddress 
         Height          =   315
         Left            =   1320
         MaxLength       =   255
         TabIndex        =   3
         Top             =   240
         Width           =   5895
      End
      Begin VB.TextBox txtUsername 
         Height          =   315
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   5
         Top             =   660
         Width           =   2595
      End
      Begin VB.Label lblPassword 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "&Password"
         Height          =   255
         Left            =   3720
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblProxyServer 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Proxy &Server"
         Height          =   255
         Left            =   60
         TabIndex        =   8
         Top             =   1140
         Width           =   1095
      End
      Begin VB.Label lblHTTPAddress 
         Alignment       =   1  'Right Justify
         Caption         =   "&HTTP Address"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label lblUserName 
         Alignment       =   1  'Right Justify
         Caption         =   "&User Name"
         Height          =   255
         Left            =   60
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame fraType 
      Caption         =   "Server Type"
      Height          =   1215
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7335
      Begin VB.TextBox txtDesc 
         BackColor       =   &H8000000F&
         Height          =   855
         Left            =   1860
         MultiLine       =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   5355
      End
      Begin VB.ListBox lstType 
         Height          =   840
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmRRConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2000-2003. All Rights Reserved
'   File:       frmRRConnection.frm
'   Author:     Toby Aldridge, November 2000
'   Purpose:    To allow the user to configure Registration Server connection
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'Revisions:
' NCJ 16 Jun 03 - Adjusted wording of m_LOCAL_DESC (MACRO 3.0 Bug 1194)
'----------------------------------------------------------------------------------------'

Option Explicit

Private Const m_SERVER_NONE = "None"
Private Const m_SERVER_LOCAL = "Local"
Private Const m_SERVER_TRIALOFFICE_DESC = "Trial Office"
Private Const m_SERVER_REMOTE = "Remote"

'description constants ( vb won't let me split them over more than one line)
' NCJ 16 Jun 03 - Adjusted wording of m_LOCAL_DESC
Private Const m_NONE_DESC = "No registration server.  This means that registration will not occur for this study."
Private Const m_LOCAL_DESC = "Registration will be carried out on the local MACRO database.  This option is only suitable for single-database studies."
Private Const m_TRIALOFFICE_DESC = "Registration will be carried out on the central MACRO server database, using the same communication settings as used for the transfer of other study data."
Private Const m_REMOTE_DESC = "Registration will be carried out on a remote server.  The communication details for the server must be entered below.  (Not yet implemented)"

Private WithEvents moRRConnection As clsRRConnection
Attribute moRRConnection.VB_VarHelpID = -1

'was ok clicked
Private mbOk As Boolean
Private mrecDescription As clsDataRecord

'true while loading controls with values
Private mbLoading As Boolean

'----------------------------------------------------------------------------------------'
Public Sub Display(lClinicalTrialId As Long, nVersionId As Integer)
'----------------------------------------------------------------------------------------'

    Set mrecDescription = RecordBuild(m_NONE_DESC, m_LOCAL_DESC, m_TRIALOFFICE_DESC, m_REMOTE_DESC)
    
    mbOk = False
    cmdOK.Enabled = False
    mbLoading = True
    'fill listbox
    Call FillServerTypeList
    
    Set moRRConnection = New clsRRConnection
    
    With moRRConnection
        Call .Load(lClinicalTrialId, nVersionId)
        Call ListCtrl_Pick(lstType, .ServerType)
        txtHTTPAddress.Text = .HTTPAddress
        txtUsername.Text = .UserName
        txtPassword.Text = .Password
        txtProxyServer.Text = .ProxyServer
        'there are no changes beacuse all properties will have been set to what they were
        Call ConnectionDetailsChange(.ConnectionDetailsChange)
    End With
    mbLoading = False
    
    Me.Icon = frmMenu.Icon
    FormCentre Me
    Me.Show vbModal
    
    If mbOk Then
        Call moRRConnection.Save
    End If
    
    Set mrecDescription = Nothing
    Set moRRConnection = Nothing

End Sub

'----------------------------------------------------------------------------------------'
Private Sub cmdCancel_Click()
'----------------------------------------------------------------------------------------'

    Unload Me
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub cmdOK_Click()
'----------------------------------------------------------------------------------------'

    mbOk = True
    Unload Me

End Sub

Private Sub Form_Load()
    Me.Icon = frmMenu.Icon
End Sub

'----------------------------------------------------------------------------------------'
Private Sub lstType_Click()
'----------------------------------------------------------------------------------------
Dim nType As Integer
    
    nType = lstType.ItemData(lstType.ListIndex)
    If Not mbLoading Then
        'only update when user changes data
        moRRConnection.ServerType = nType
    End If
    txtDesc.Text = mrecDescription(nType + 1)
    
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub moRRConnection_ConnectionDetailsChange(bAllow As Boolean)
'----------------------------------------------------------------------------------------'

    Call ConnectionDetailsChange(bAllow)
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub moRRConnection_HasChanges(bHasChanges As Boolean)
'----------------------------------------------------------------------------------------'

    cmdOK.Enabled = (moRRConnection.IsValid And bHasChanges)

End Sub

'----------------------------------------------------------------------------------------'
Private Sub moRRConnection_IsValid(bValid As Boolean)
'----------------------------------------------------------------------------------------'

        cmdOK.Enabled = (bValid And moRRConnection.HasChanges)

End Sub

'----------------------------------------------------------------------------------------'
Private Sub moRRConnection_ValidHTTPAddress(bIsValid As Boolean)
'----------------------------------------------------------------------------------------'

    Call ColourControl(txtHTTPAddress, bIsValid)

End Sub

'----------------------------------------------------------------------------------------'
Private Sub moRRConnection_ValidPassword(bIsValid As Boolean)
'----------------------------------------------------------------------------------------'

    Call ColourControl(txtPassword, bIsValid)

End Sub

'----------------------------------------------------------------------------------------'
Private Sub moRRConnection_ValidProxyServer(bIsValid As Boolean)
'----------------------------------------------------------------------------------------'

    Call ColourControl(txtProxyServer, bIsValid)

End Sub

'----------------------------------------------------------------------------------------'
Private Sub moRRConnection_ValidUserName(bIsValid As Boolean)
'----------------------------------------------------------------------------------------'

    Call ColourControl(txtUsername, bIsValid)
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub txtHTTPAddress_Change()
'----------------------------------------------------------------------------------------'
    
    If Not mbLoading Then
        'only update when user changes data
        moRRConnection.HTTPAddress = txtHTTPAddress.Text
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub txtPassword_Change()
'----------------------------------------------------------------------------------------'

    If Not mbLoading Then
        'only update when user changes data
        moRRConnection.Password = txtPassword.Text
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub txtProxyServer_Change()
'----------------------------------------------------------------------------------------'

    If Not mbLoading Then
        'only update when user changes data
        moRRConnection.ProxyServer = txtProxyServer.Text
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub txtUsername_Change()
'----------------------------------------------------------------------------------------'

    If Not mbLoading Then
        'only update when user changes data
        moRRConnection.UserName = txtUsername.Text
    End If

End Sub

'----------------------------------------------------------------------------------------'
Private Sub FillServerTypeList()
'----------------------------------------------------------------------------------------'

    With lstType
        .Clear
        .AddItem m_SERVER_NONE
        .ItemData(lstType.NewIndex) = eRRServerType.RRNone
        .AddItem m_SERVER_LOCAL
        .ItemData(lstType.NewIndex) = eRRServerType.RRLocal
        .AddItem m_SERVER_TRIALOFFICE_DESC
        .ItemData(lstType.NewIndex) = eRRServerType.RRTrialOffice
        .AddItem m_SERVER_REMOTE
        .ItemData(lstType.NewIndex) = eRRServerType.RRRemote
    End With
    
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub ConnectionDetailsChange(bAllow As Boolean)
'----------------------------------------------------------------------------------------'

    fraConnection.Enabled = bAllow
    lblHTTPAddress.Enabled = bAllow
    lblUserName.Enabled = bAllow
    lblPassword.Enabled = bAllow
    lblProxyServer.Enabled = bAllow
    txtHTTPAddress.Enabled = bAllow
    txtUsername.Enabled = bAllow
    txtPassword.Enabled = bAllow
    txtProxyServer.Enabled = bAllow

End Sub
