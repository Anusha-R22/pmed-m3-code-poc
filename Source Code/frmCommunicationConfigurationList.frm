VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCommunicationConfigurationList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Communication Configuration Settings"
   ClientHeight    =   5505
   ClientLeft      =   2370
   ClientTop       =   1830
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   5475
   Begin VB.Frame fmeSettings 
      Caption         =   "Settings"
      Height          =   4575
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   375
         Left            =   3240
         TabIndex        =   1
         Top             =   4080
         Width           =   855
      End
      Begin VB.CommandButton cmdProperties 
         Caption         =   "&Properties"
         Height          =   375
         Left            =   4320
         TabIndex        =   2
         Top             =   4080
         Width           =   855
      End
      Begin MSComctlLib.ListView lvwSettings 
         Height          =   3735
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Study Office"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Site"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Effective From"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Effective To"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   4920
      Width           =   975
   End
End
Attribute VB_Name = "frmCommunicationConfigurationList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       frmCommunicationConfigurationList.frm
'   Author:     Paul Norris 22/07/99
'   Purpose:    Viewing window for the communication settings
'               stored in the TrialOffice table in Macro.mdb
'---------------------------------------------------------------------
'
'---------------------------------------------------------------------
'   Revisions:
'   PN  20/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   PN  30/09/99    Changed class names
'                   clsCommunicationData to clsCommunication
'                   because prog id is too long with original name
'   WillC 11/11/99    Added Error handlers
'   Mo Morris   17/11/99    DAO to ADO conversion
'  WillC 11 / 12 / 99
'          Changed the Following where present from Integer to Long  ClinicalTrialId
'           CRFPageId,VisitId,CRFElementID
'   NCJ 16/12/99    Only enable New button if user is authorised to change settings
'   NCJ 7/3/00 - SRs2780,3110 Major rewriting using new clsComms class
'   TA 25/04/2000   subclassing removed
'---------------------------------------------------------------------

Option Explicit

Private mbIsNew As Boolean
' NCJ 7/3/00
Private mcolCommConfigs As clsComms

'---------------------------------------------------------------------
Private Sub cmdClose_Click()
'---------------------------------------------------------------------

    Unload Me
    
End Sub

'---------------------------------------------------------------------
Private Sub EnableProperties(bEnable As Boolean)
'---------------------------------------------------------------------

    cmdProperties.Enabled = bEnable

End Sub

'---------------------------------------------------------------------
Private Sub cmdProperties_Click()
'---------------------------------------------------------------------
    
    mbIsNew = False
    Call EditSetting

End Sub

'---------------------------------------------------------------------
Private Sub cmdNew_Click()
'---------------------------------------------------------------------

    Set lvwSettings.SelectedItem = Nothing
    mbIsNew = True
    Call EditSetting

End Sub


'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
   
   Me.Icon = frmMenu.Icon
   
    Call LoadSettings
    Me.BackColor = glFormColour
    
    ' NCJ 16 Dec 99 - Only enable New if the user has access rights
    If goUser.CheckPermission(gsFnChangeCommSettings) Then
        cmdNew.Enabled = True
    Else
        cmdNew.Enabled = False
    End If
    
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

'---------------------------------------------------------------------
Private Sub EditSetting()
'---------------------------------------------------------------------
' Edit the selected setting
'---------------------------------------------------------------------
Dim oSetting As clsCommunication
    
    On Error GoTo ErrHandler
        
    With lvwSettings
        If .SelectedItem Is Nothing Then
            ' New comm object
            Set oSetting = New clsCommunication
        
        ElseIf IsSelectedItemEditable Then
            ' Key of object is stored in item Tag
            Set oSetting = mcolCommConfigs.Item(.SelectedItem.Tag)
        
        End If
            
        If Not oSetting Is Nothing Then
        
            With frmCommunicationConfiguration
                .IsNew = mbIsNew 'REM 28/05/03 - set to tell if creating new setting or editing old one
                Set .Component = oSetting
                Set .CommConfigs = mcolCommConfigs
                .Show vbModal
            End With
    
            Set oSetting = Nothing

            ' Refresh collection and listview
            ' (which will instate any new object)
            Call LoadSettings

        End If
    End With

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "EditSetting")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
   
End Sub

'---------------------------------------------------------------------
Private Sub LoadSettings()
'---------------------------------------------------------------------
' Rewritten, NCJ 7/3/00, to use new clsComms collection class
'---------------------------------------------------------------------
'Dim rsSettings As ADODB.Recordset
'Dim sSQL As String
Dim olistItem As ListItem
Dim iSelectedIndex As Integer
Dim oCommConfig As clsCommunication

   On Error GoTo ErrHandler
   
   Set mcolCommConfigs = New clsComms
   
   ' Create the collection from the database
   Call mcolCommConfigs.Load
   
'    ' read a sub-set of the available config settings
'    sSQL = "select trialoffice, site, effectivefrom, effectiveto from trialoffice"
'    Set rsSettings = New ADODB.Recordset
'    rsSettings.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    iSelectedIndex = -1
    With lvwSettings
        If Not .SelectedItem Is Nothing Then
            iSelectedIndex = .SelectedItem.Index
        End If
        .ListItems.Clear
    End With
    
    ' Load the settings into the listview
    For Each oCommConfig In mcolCommConfigs
        Set olistItem = lvwSettings.ListItems.Add(, , oCommConfig.TrialOffice)
         olistItem.SubItems(1) = oCommConfig.Site
         olistItem.SubItems(2) = oCommConfig.EffectiveFrom
         olistItem.SubItems(3) = oCommConfig.EffectiveTo
         ' Store key in Tag field
         olistItem.Tag = oCommConfig.CommKey
    Next
    
'    With rsSettings
'        Do While Not .EOF
'            Set oListItem = lvwSettings.ListItems.Add(, , .Fields("trialOffice"))
'             oListItem.SubItems(1) = IIf(IsNull(.Fields("site")), vbNullString, .Fields("site"))
'             oListItem.SubItems(2) = Format(CDate(.Fields("effectivefrom")), frmMenu.DefaultDateFormat)
'             oListItem.SubItems(3) = Format(CDate(.Fields("effectiveto")), frmMenu.DefaultDateFormat)
'            .MoveNext
'        Loop
'    End With
'    rsSettings.Close
'    Set rsSettings = Nothing
    
    ' set all column widths to autoresize
    Call LVSetAllColWidths(lvwSettings, LVSCW_AUTOSIZE_USEHEADER)
    Call LVSetStyleEx(lvwSettings, LVSTHeaderDragDrop Or LVSTFullRowSelect, True)
    
    If iSelectedIndex > 0 Then
        lvwSettings.ListItems(iSelectedIndex).Selected = True
    End If
        
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "LoadSettings")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
   

End Sub

'---------------------------------------------------------------------
Private Sub lvwSettings_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'---------------------------------------------------------------------

    With ColumnHeader
        Call SortList(.Index - 1, Abs(lvwSettings.SortOrder - 1))
    End With

End Sub

'---------------------------------------------------------------------
Private Sub SortList(iSortKey As Integer, iSortOrder As Integer)
'---------------------------------------------------------------------
    
   On Error GoTo ErrHandler
   
    With lvwSettings
        .SortKey = iSortKey
        .SortOrder = iSortOrder
        
        Select Case .ColumnHeaders(.SortKey + 1).Text
        Case "Effective To", "Effective From"
            Call SortListview(Me.lvwSettings, iSortKey, .SortOrder, LVTDate)
        Case Else
            .Sorted = True
        End Select
    End With
        
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "SortList")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
   
End Sub

'---------------------------------------------------------------------
Private Sub lvwSettings_DblClick()
'---------------------------------------------------------------------

    Call EditSetting

End Sub

'---------------------------------------------------------------------
Private Function IsSelectedItemEditable() As Boolean
'---------------------------------------------------------------------

    On Error GoTo InvalidDate
    
    If lvwSettings.SelectedItem Is Nothing Then
        IsSelectedItemEditable = False
    Else
'   ATN 17/12/99
'   Don't do this check here, becuase then the user doesn't know why they can't edit the item
        ' if the setting has an effectiveto that is older than today then
        ' the setting can not be edited
        If lvwSettings.SelectedItem.SubItems(3) <> vbNullString Then
            If CDate(lvwSettings.SelectedItem.SubItems(3)) < Now Then
                MsgBox "This setting has expired and cannot be edited", vbOKOnly + vbInformation + vbApplicationModal
                IsSelectedItemEditable = False
            Else
                IsSelectedItemEditable = True
            End If
        Else
            IsSelectedItemEditable = True
        End If
    End If
    
InvalidDate:
    
End Function

'---------------------------------------------------------------------
Private Sub lvwSettings_ItemClick(ByVal Item As MSComctlLib.ListItem)
'---------------------------------------------------------------------

    Call EnableProperties(True)

End Sub
