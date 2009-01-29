VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReferences 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "References"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   2565
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   1500
      TabIndex        =   4
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRemoveTrialDocument 
      Caption         =   "Remove"
      Height          =   375
      Left            =   1500
      TabIndex        =   1
      Top             =   4260
      Width           =   975
   End
   Begin VB.CommandButton cmdAddTrialDocument 
      Caption         =   "Add..."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4260
      Width           =   975
   End
   Begin MSComctlLib.ListView lsvTrialDocuments 
      Height          =   4095
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   7223
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imglistSmallIcons 
      Left            =   735
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imglistLargeIcons 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
End
Attribute VB_Name = "frmReferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998-2006. All Rights Reserved
'   File:       frmReferences.frm
'   Author:     Andrew Newbigging, March 98
'   Purpose:    Displays list of attached reference documents and allows user to
'   add, remove or show the documents.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
 '  1   Andrew Newbigging       24/03/98
'   2   Joanne Lau              14/05/98
'   3   Joanne Lau              15/05/98
'   4   Andrew Newbigging       15/05/98
'   5   Andrew Newbigging       15/05/98
'       Mo Morris               20/11/98
'       SPR 553
'       cmdAddTrialDocument_Click chenge to give a better error message if FileCopy of
'       trial documents fails.
'   PN  20/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   NCJ 1 Oct 99 - Use gsDOCUMENTS_PATH
'   WillC   10/11/99 Added the error handlers
'   Mo Morris   15/11/99    DAO to ADO conversion
' NCJ 10/12/99 - Added user access rights check
'   TA 29/04/2000   subclassing removed
'   TA 20/6/2000 SR3632: doevents after open dialog to stop ghosting and hourglass functions during copying
' NCJ 24/10/00 (Related to SR3559) Get confirmation before deleting reference
'   NCJ 27 Oct 00 - Check reference doesn't already exist before adding
'   NCJ 11 May 06 - Check study access mode
'------------------------------------------------------------------------------------'

Option Explicit
Option Base 0
Option Compare Binary


'---------------------------------------------------------------------
Private Sub Form_Activate()
'---------------------------------------------------------------------
'   ATN 17/12/99    SR 1763
'   Need to reset the selected item
  On Error GoTo ErrHandler
    
    frmMenu.ChangeSelectedItem "", ""

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "Form_Activate")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'---------------------------------------------------------------------
Private Sub cmdClose_Click()
'---------------------------------------------------------------------

    Unload Me

End Sub

'---------------------------------------------------------------------
Private Sub cmdAddTrialDocument_Click()
'---------------------------------------------------------------------
'Adds references to the sdd
' Use gsDOCUMENTS_PATH - NCJ 1 Oct 99
' NCJ 27/10/00 - Check if document already exists
'---------------------------------------------------------------------
Dim itmX As ListItem
Dim sDestFile As String
Dim sErrMsg As String
Dim sDocument As String

    On Error Resume Next
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    CommonDialog1.ShowOpen
    
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        On Error GoTo 0
    End If
    
    sDocument = CommonDialog1.FileTitle
    
    ' NCJ 27/10/00
    If TrialDocumentExists(frmMenu.ClinicalTrialId, frmMenu.VersionId, sDocument) Then
        Call DialogError("The reference '" & sDocument & "' already exists")
        Exit Sub
    End If
    
    'prepare to copy document into the documents directory
    ' NB FileTitle is file name excluding path
    sDestFile = gsDOCUMENTS_PATH & sDocument
    'TA 20/6/2000 SR3632: doevents to stop ghosting
    DoEvents
    On Error Resume Next
    HourglassOn
    FileCopy CommonDialog1.FileName, sDestFile
    HourglassOff
    If Err.Number > 0 Then
        sErrMsg = "Copying " & CommonDialog1.FileName _
                    & " to " & sDestFile & " failed. " & vbCrLf
        sErrMsg = sErrMsg & " Error number " & Err.Number & vbCrLf _
                    & Err.Description
        MsgBox sErrMsg
        Exit Sub
    Else
        On Error GoTo 0
    End If
    'add new document into ListView
    Set itmX = lsvTrialDocuments.ListItems.Add(, , sDocument, gsDOCUMENT_LABEL, gsDOCUMENT_LABEL)
    itmX.SubItems(1) = sDestFile
    'Save name of pdf document to db (studydocument table)
    gdsAddTrialDocument frmMenu.ClinicalTrialId, frmMenu.VersionId, sDocument 'changed from CD1.filename
    
    lsvTrialDocuments.Refresh

End Sub

'---------------------------------------------------------------------
Private Sub cmdShow_Click()
'---------------------------------------------------------------------
' Show the selected document
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If Not (lsvTrialDocuments.SelectedItem Is Nothing) Then
    
        ShowDocument Me.hWnd, lsvTrialDocuments.SelectedItem.SubItems(1)
    
    End If
  
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdShow_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------
' NCJ 10/12/99 - Added user access rights check
' NCJ 11 May 06 - Check study access mode
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    Me.Icon = frmMenu.Icon
    
    ' Disable Remove & Show buttons until something is selected
    cmdRemoveTrialDocument.Enabled = False
    cmdShow.Enabled = False
        
    ' Only enable Add button if they have the right
    cmdAddTrialDocument.Enabled = (goUser.CheckPermission(gsFnAttachRefDoc) _
                                    And frmMenu.StudyAccessMode >= sdReadWrite)
'    If goUser.CheckPermission(gsFnAttachRefDoc) Then
'        cmdAddTrialDocument.Enabled = True
'    Else
'        cmdAddTrialDocument.Visible = False
'    End If
    
    Me.Left = 0
    Me.Top = frmMenu.Height / 6
    
    Call RefreshTrialDocumentList
    
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

'---------------------------------------------------------------------
Private Sub lsvTrialDocuments_DblClick()
'---------------------------------------------------------------------
' Show the selected document (same as cmdShow)
'---------------------------------------------------------------------

    Call cmdShow_Click
    
End Sub

'---------------------------------------------------------------------
Private Sub RefreshTrialDocumentList()
'---------------------------------------------------------------------
' Create an object variable for the ColumnHeader object.
'---------------------------------------------------------------------
Dim clmX As ColumnHeader
Dim imgX As ListImage

    On Error GoTo ErrHandler
    
    ' Add ColumnHeaders with appropriate widths
    Set clmX = lsvTrialDocuments.ColumnHeaders.Add(, , "Name", 1700)
    Set clmX = lsvTrialDocuments.ColumnHeaders.Add(, , "Path", 1700)
    
    lsvTrialDocuments.View = lvwSmallIcon ' Set View property to Icon.
    lsvTrialDocuments.Arrange = lvwAutoLeft
    
    Set imgX = imglistLargeIcons.ListImages.Add(, gsDOCUMENT_LABEL, LoadResPicture(gsDOCUMENT_LABEL, vbResIcon))
    Set imgX = imglistSmallIcons.ListImages.Add(, gsDOCUMENT_LABEL, LoadResPicture(gsDOCUMENT_LABEL, vbResIcon))
    
    ' Set Icons property
    lsvTrialDocuments.Icons = imglistLargeIcons
    lsvTrialDocuments.SmallIcons = imglistSmallIcons
    
    ' Refresh the data in the list
    RefreshTrialDocuments
 
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "RefreshTrialDocumentList")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub RefreshTrialDocuments()
'---------------------------------------------------------------------
' Create a variable to add ListItem objects and receive the list of trial documents
'---------------------------------------------------------------------
Dim itmX As ListItem
Dim rsTemp As ADODB.Recordset
Dim tmpfilename As String
Dim tmpX As Integer

    On Error GoTo ErrHandler
    
    ' Get the list of trials
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gdsTrialDocumentList(frmMenu.ClinicalTrialId, frmMenu.VersionId)
    
    ' While the record is not the last record, add a ListItem object.
    While Not rsTemp.EOF
    
        ' Use gsDOCUMENTS_PATH - NCJ 1/10/99
        tmpfilename = gsDOCUMENTS_PATH & rsTemp!DocumentPath
    
        Set itmX = lsvTrialDocuments.ListItems.Add(, , rsTemp!DocumentPath, gsDOCUMENT_LABEL, gsDOCUMENT_LABEL)
        itmX.SubItems(1) = tmpfilename
        rsTemp.MoveNext   ' Move to next record.
    Wend
    rsTemp.Close
    Set rsTemp = Nothing
    
    cmdRemoveTrialDocument.Enabled = False
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "RefreshTrialDocuments")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdRemoveTrialDocument_Click()
'---------------------------------------------------------------------
' NCJ 24/10/00 (Related to SR3559) Get confirmation before deleting reference
'---------------------------------------------------------------------
Dim sRefDoc As String

    On Error GoTo ErrHandler
    
    If Not (lsvTrialDocuments.SelectedItem Is Nothing) Then
        sRefDoc = lsvTrialDocuments.SelectedItem.Text
        If DialogQuestion("Are you sure you wish to remove the reference '" & sRefDoc & "' ?") = vbYes Then
            gdsDeleteTrialDocument frmMenu.ClinicalTrialId, frmMenu.VersionId, sRefDoc
            lsvTrialDocuments.ListItems.Remove lsvTrialDocuments.SelectedItem.Index
            cmdRemoveTrialDocument.Enabled = False
            
            lsvTrialDocuments.Refresh
        End If
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdRemoveTrialDocument_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub lsvTrialDocuments_ItemClick(ByVal Item As MSComctlLib.ListItem)
'---------------------------------------------------------------------
' NCJ 11 May 06 - Include Study Access mode in enablability
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    cmdRemoveTrialDocument.Enabled = goUser.CheckPermission(gsFnRemoveRefDoc) _
                                And (frmMenu.StudyAccessMode >= sdReadWrite)
'    If goUser.CheckPermission(gsFnRemoveRefDoc) Then
'        cmdRemoveTrialDocument.Enabled = True
'    Else
'        cmdRemoveTrialDocument.Enabled = False
'    End If
    cmdShow.Enabled = True

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "lsvTrialDocuments_ItemClick")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub lsvTrialDocuments_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    Select Case KeyCode
    Case vbKeyReturn
        Call lsvTrialDocuments_DblClick
    Case vbKeyEscape
        Unload Me
    End Select
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "lsvTrialDocuments_KeyUp")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'---------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'---------------------------------------------------------------------

    frmMenu.HideReferences

End Sub

