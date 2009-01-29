VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOC 
   Caption         =   "OC Discrepancies"
   ClientHeight    =   3300
   ClientLeft      =   7500
   ClientTop       =   6795
   ClientWidth     =   6630
   Icon            =   "frmOCDiscrepancies.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   6630
   Begin MSComctlLib.ImageList imgDiscs 
      Left            =   6000
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOCDiscrepancies.frx":08CA
            Key             =   "done"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOCDiscrepancies.frx":0E5C
            Key             =   "selected"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwDiscs 
      Height          =   1155
      Left            =   180
      TabIndex        =   3
      Top             =   1080
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2037
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame fraScope 
      Caption         =   "Discrepancy Scope"
      Height          =   735
      Left            =   60
      TabIndex        =   7
      Top             =   2460
      Width           =   6495
      Begin VB.CommandButton cmdOpen 
         Caption         =   "&Schedule"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5100
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtSubject 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   300
         Width           =   975
      End
      Begin VB.TextBox txtSite 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   300
         Width           =   1035
      End
      Begin VB.TextBox txtStudy 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   660
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   300
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Go"
         Height          =   375
         Left            =   5100
         TabIndex        =   8
         Top             =   2500
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Subject"
         Height          =   255
         Left            =   3300
         TabIndex        =   13
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Site"
         Height          =   255
         Left            =   1740
         TabIndex        =   11
         Top             =   300
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Study"
         Height          =   255
         Left            =   60
         TabIndex        =   9
         Top             =   300
         Width           =   495
      End
   End
   Begin VB.Frame fraSite 
      Caption         =   "Search Criteria"
      Height          =   735
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   6495
      Begin VB.ComboBox cboSite 
         Height          =   315
         Left            =   2340
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   1215
      End
      Begin VB.ComboBox cboStudy 
         Height          =   315
         Left            =   660
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   1215
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Search"
         Height          =   375
         Left            =   5160
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Study"
         Height          =   255
         Left            =   60
         TabIndex        =   16
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Site"
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   300
         Width           =   315
      End
   End
   Begin VB.Frame fraDiscs 
      Caption         =   "set at runtime"
      Height          =   1515
      Left            =   60
      TabIndex        =   6
      Top             =   840
      Width           =   6495
   End
End
Attribute VB_Name = "frmOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       frmOC.frm
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   Author:     Toby Aldridge, March 2002
'   Purpose:    User form for automating cutting and pasting of
'                   discrepancies from OC
'----------------------------------------------------------------------------------------'
' Revisions:
'ta 27/05/2002 CBB 2.2.13.46: combos now have an all studies/sites option
'----------------------------------------------------------------------------------------'

Option Explicit

Private msCon As String
Private moUser As MACROUser

'column enumerations
Private Enum eViewCol
    vcDocNo = 0
    vcOCId = 1
    vcDiscText = 2
    vcStudy = 3
    vcSite = 4
    vcSubject = 5
    vcMax = 5
End Enum

'view name
Private Const ms_VIEW_NAME = "OCDISCREPANCIES"

'column names
Private Const ms_COL_DOCNO = "DOCNO"
Private Const ms_COL_OCID = "OCID"
Private Const ms_COL_DISCTEXT = "DISCTEXT"
Private Const ms_COL_STUDY = "STUDY"
Private Const ms_COL_SITE = "SITE"
Private Const ms_COL_SUBJECT = "SUBJECT"

'listview column header list - any extra columns in the view will have the column
'   name as the listview column title
Const m_HEADER_LIST = "Document No,OC Id,Comment,Study,Site,Subject"

'button to go to schedule clicked
Public Event OpenSubject(sStudy As String, sSite As String, sSubjectLabel As String, lOCId As Long, sDiscText As String)

'OC row selected
Public Event OCSelected(lOCId As Long, sDiscText As String)

'form is closed
Public Event Closed()


'----------------------------------------------------------------------------------------'
Private Sub cmdOpen_Click()
'----------------------------------------------------------------------------------------'
'send message to open subject
Dim oItem As ListItem

    'clear any old ones
    For Each oItem In lvwDiscs.ListItems
        If oItem.SmallIcon = "selected" Then
            oItem.SmallIcon = Empty
            'the next two lines are the only way I could get the listview to refresh the removed icon
            DoEvents
            lvwDiscs.Refresh
        End If
    Next
    With lvwDiscs.SelectedItem
        RaiseEvent OpenSubject(txtStudy.Text, txtSite.Text, txtSubject.Text, _
                        .SubItems(eViewCol.vcOCId), _
                        .Text & " - " & .SubItems(eViewCol.vcDiscText))
        .SmallIcon = "selected"
    End With
    
End Sub

'------------------------------------------------------------------------------------'
Private Sub cmdRefresh_Click()
'------------------------------------------------------------------------------------'
'fill listview with OC Discrepancies
'------------------------------------------------------------------------------------'

Dim oCon As Connection
Dim rs As Recordset
Dim sSQL As String
Dim vHeaders As Variant
Dim vData As Variant
Dim bRecordsFound As Boolean
Dim sHeaders As String
Dim i As Long

    On Error GoTo ErrLabel
    
    Screen.MousePointer = vbHourglass
    
    lvwDiscs.ListItems.Clear
    
    sSQL = "select * from " & ms_VIEW_NAME
    sSQL = sSQL & " where " & ms_COL_OCID & " not in "
    'sSQL = sSQL & "(0,1)"
    'exclude OC discs that already have their id in our MIMessage table
    sSQL = sSQL & "(select distinct MIMESSAGEOCDISCREPANCYID from MIMESSAGE)"
    
    'ta 27/05/2002 CBB 2.2.13.46: use the all studies/sites option
    If cboStudy.Text <> "All studies" Then
        'filter on study
        sSQL = sSQL & " AND " & ms_COL_STUDY & "='" & cboStudy.Text & "'"
    End If

    If cboSite.Text <> "All sites" Then
        'filter on site
       sSQL = sSQL & " AND " & ms_COL_SITE & "='" & cboSite.Text & "'"
    End If
    
    Set oCon = New Connection
    oCon.Open msCon
    oCon.CursorLocation = adUseClient
    
    Set rs = New Recordset
    'open read-only recordset that allows a recordcount
    rs.Open sSQL, oCon, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rs.RecordCount < 1 Then
        'no matching records
        bRecordsFound = False
        fraDiscs.Caption = "OC Discrepancies (none found)"
    Else
        bRecordsFound = True
        vData = rs.GetRows
        fraDiscs.Caption = "OC Discrepancies (" & UBound(vData, 2) + 1 & ")"
        
        'display extra columns
        sHeaders = m_HEADER_LIST
        For i = (eViewCol.vcMax + 1) To rs.Fields.Count - 1
            sHeaders = sHeaders & "," & rs.Fields(i).Name
        Next
    End If
    
    rs.Close
    oCon.Close
    Set rs = Nothing
    Set oCon = Nothing
    
    cmdOpen.Enabled = bRecordsFound
    
    If bRecordsFound Then
        vHeaders = Split(sHeaders, ",")
        lvw_FromArray lvwDiscs, vData, vHeaders
        'adjust width to allow icon
        lvwDiscs.ColumnHeaders(1).Width = lvwDiscs.ColumnHeaders(1).Width + ((imgDiscs.ImageWidth + 6) * Screen.TwipsPerPixelX)
        'simulating clicking the first item
        lvwDiscs_ItemClick lvwDiscs.SelectedItem
    End If

    Screen.MousePointer = vbNormal
    Exit Sub
    
ErrLabel:
    Screen.MousePointer = vbNormal
    MsgBox "Unable to retrieve discrepancies from Oracle Clinical", vbCritical, "OC Discrepnacies"
    
End Sub


'----------------------------------------------------------------------------------------'
Private Sub Form_Load()
'----------------------------------------------------------------------------------------'
    lvwDiscs.SmallIcons = imgDiscs
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Form_Resize()
'----------------------------------------------------------------------------------------'
On Error GoTo ErrorLabel
    fraDiscs.Width = Me.ScaleWidth - 120
    lvwDiscs.Width = fraDiscs.Width - 240
    
    fraDiscs.Height = Me.ScaleHeight - fraSite.Height - fraScope.Height - 240
    lvwDiscs.Height = fraDiscs.Height - 360
    fraScope.Top = fraDiscs.Top + fraDiscs.Height + 60

ErrorLabel:
End Sub

'------------------------------------------------------------------------------------'
Friend Function Display(sCon As String, oUser As MACROUser)
'------------------------------------------------------------------------------------'

'------------------------------------------------------------------------------------'
    'display at last used size
    SetFormDimensions
    msCon = sCon
    Set moUser = oUser
    
    'load up study/site combos
    LoadCombos
    Me.Show
End Function

'------------------------------------------------------------------------------------'
Private Sub Form_Unload(Cancel As Integer)
'------------------------------------------------------------------------------------'

    'tell listeners that form is closing
    RaiseEvent Closed
    SaveFormDimensions
    
End Sub

'------------------------------------------------------------------------------------'
Private Sub lvwDiscs_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'------------------------------------------------------------------------------------'
'sort columns
'------------------------------------------------------------------------------------'

    lvw_Sort lvwDiscs, ColumnHeader

End Sub

'------------------------------------------------------------------------------------'
Private Sub lvwDiscs_ItemClick(ByVal Item As MSComctlLib.ListItem)
'------------------------------------------------------------------------------------'

    'set chosen subject identifiers
    txtStudy.Text = Item.SubItems(eViewCol.vcStudy)
    txtSite.Text = Item.SubItems(eViewCol.vcSite)
    txtSubject.Text = Item.SubItems(eViewCol.vcSubject)
    'inform listeners that something has been selected
    RaiseEvent OCSelected(Item.SubItems(eViewCol.vcOCId), _
        Item.Text & " - " & Item.SubItems(eViewCol.vcDiscText))
End Sub

'------------------------------------------------------------------------------------'
Friend Sub MarkAsRaised(lOCId As Long)
'------------------------------------------------------------------------------------'
'sets the 'discrepancy raised' icon in the row with the matching OC ID
' to show that the discrepancy has been raised in MACRO
'------------------------------------------------------------------------------------'
Dim oItem As ListItem

    For Each oItem In lvwDiscs.ListItems
        If CLng(oItem.SubItems(eViewCol.vcOCId)) = lOCId Then
            oItem.SmallIcon = "done"
            Exit For
        End If
    Next
End Sub

'------------------------------------------------------------------------------------'
Private Sub LoadCombos()
'------------------------------------------------------------------------------------'
'load study site combos
'------------------------------------------------------------------------------------'
Dim i As Long

    cboStudy.Clear

    'ta 27/05/2002 CBB 2.2.13.46: give them an all studies/sites option
    'cboStudy.AddItem "All studies"
    
    For i = 1 To moUser.GetAllStudies.Count
        cboStudy.AddItem moUser.GetAllStudies(i).StudyName
        cboStudy.ItemData(cboStudy.NewIndex) = moUser.GetAllStudies(i).StudyId
    Next
    
    cboStudy.ListIndex = 0
    
End Sub

'------------------------------------------------------------------------------------'
Private Sub cboStudy_Click()
'------------------------------------------------------------------------------------'
'load site combo
'------------------------------------------------------------------------------------'
Dim i As Long

    cboSite.Clear
    cboSite.AddItem "All sites"
    
    For i = 1 To moUser.GetAllSites(cboStudy.ItemData(cboStudy.ListIndex)).Count
        cboSite.AddItem moUser.GetAllSites(cboStudy.ItemData(cboStudy.ListIndex))(i).Site
    Next
    
    cboSite.ListIndex = 0
End Sub
'----------------------------------------------------------------------------------------'
Private Sub SaveFormDimensions()
'----------------------------------------------------------------------------------------'
' Saves a forms's dimension to the registry
'   must go  in QueryUnload event of form
'----------------------------------------------------------------------------------------'
 Dim sSetting As String
 Dim nWindowState As Integer

    With Me
        If .WindowState = vbMinimized Then
            .WindowState = vbNormal
        End If
        nWindowState = .WindowState
        If nWindowState <> vbNormal Then
            'if minimised or maximised restore
            .WindowState = vbNormal
        End If
        
        sSetting = .Top & "," & .Left & "," & .Height & "," & .Width & "," & nWindowState
        
        Call SaveSetting(App.Title, "Form Dimensions", Mid(Me.Name, 4), sSetting)
    
    End With

End Sub

'----------------------------------------------------------------------------------------'
Private Sub SetFormDimensions()
'----------------------------------------------------------------------------------------'
' set a form's dimensions according to registry settings
'   call before using (form).show vbModal
'----------------------------------------------------------------------------------------'
Dim sSetting As String

    sSetting = GetSetting(App.Title, "Form Dimensions", Mid(Me.Name, 4), "")

    If sSetting = "" Then
        'not found
        Call FormCentre(Me)
    Else
        With Me
            .Top = Split(sSetting, ",")(0)
            .Left = Split(sSetting, ",")(1)
            .Height = Split(sSetting, ",")(2)
            .Width = Split(sSetting, ",")(3)
            .WindowState = Split(sSetting, ",")(4)
        End With
    End If

End Sub
