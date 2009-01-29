VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmViewDiscrepanciesModal 
   ClientHeight    =   8775
   ClientLeft      =   2655
   ClientTop       =   2205
   ClientWidth     =   13170
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   13170
   Begin VB.CommandButton cmdShowDetails 
      Caption         =   "S&how Details..."
      Height          =   375
      Left            =   4680
      TabIndex        =   33
      Top             =   7800
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwQuestions 
      Height          =   3660
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   6456
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdPlanned 
      Caption         =   "&Planned..."
      Height          =   375
      Left            =   6000
      TabIndex        =   31
      Top             =   7740
      Width           =   1215
   End
   Begin VB.CommandButton cmdCloseForm 
      Caption         =   "&Close"
      Height          =   375
      Left            =   11760
      TabIndex        =   30
      Top             =   8340
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4560
      Top             =   8280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPrintDCF 
      Caption         =   "P&rint Data Clarification Forms"
      Height          =   375
      Left            =   1700
      MousePointer    =   1  'Arrow
      TabIndex        =   29
      Top             =   8340
      Width           =   2500
   End
   Begin VB.CommandButton cmdPrintListing 
      Caption         =   "Print &Listing"
      Height          =   375
      Left            =   100
      MousePointer    =   1  'Arrow
      TabIndex        =   28
      Top             =   8340
      Width           =   1500
   End
   Begin VB.PictureBox picScope 
      Height          =   3675
      Left            =   240
      ScaleHeight     =   3615
      ScaleWidth      =   3555
      TabIndex        =   17
      Top             =   4440
      Width           =   3615
      Begin VB.TextBox txtSubjectLabel 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   840
         Width           =   1515
      End
      Begin VB.TextBox txtVisit 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2595
      End
      Begin VB.TextBox txteForm 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2595
      End
      Begin VB.TextBox txtStudy 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   2595
      End
      Begin VB.TextBox txtSite 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   480
         Width           =   2595
      End
      Begin VB.TextBox txtPerson 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtQuestion 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1920
         Width           =   2595
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1515
      End
      Begin VB.Label lblSubjectLabel 
         Caption         =   "Subject"
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   27
         Top             =   840
         Width           =   600
      End
      Begin VB.Label lblPriority 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3060
         TabIndex        =   26
         Top             =   2280
         Width           =   315
      End
      Begin VB.Label lblEform 
         Caption         =   "eForm"
         Height          =   255
         Index           =   7
         Left            =   60
         TabIndex        =   25
         Top             =   1560
         Width           =   480
      End
      Begin VB.Label lblVisit 
         Caption         =   "Visit"
         Height          =   255
         Index           =   6
         Left            =   60
         TabIndex        =   24
         Top             =   1200
         Width           =   300
      End
      Begin VB.Label lblStudy 
         Caption         =   "Study"
         Height          =   255
         Left            =   60
         TabIndex        =   23
         Top             =   120
         Width           =   420
      End
      Begin VB.Label lblSite 
         Caption         =   "Site"
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   22
         Top             =   480
         Width           =   300
      End
      Begin VB.Label lblPerson 
         Caption         =   "Id"
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   21
         Top             =   840
         Width           =   240
      End
      Begin VB.Label lblQuestion 
         Caption         =   "Question"
         Height          =   255
         Index           =   3
         Left            =   60
         TabIndex        =   20
         Top             =   1920
         Width           =   660
      End
      Begin VB.Label lblPriority1 
         Caption         =   "Priority"
         Height          =   255
         Left            =   2460
         TabIndex        =   19
         Top             =   2280
         Width           =   525
      End
      Begin VB.Label lblStatus 
         Caption         =   "Status"
         Height          =   255
         Left            =   60
         TabIndex        =   18
         Top             =   2280
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit..."
      Height          =   375
      Left            =   11760
      TabIndex        =   13
      Top             =   7740
      Width           =   1215
   End
   Begin VB.CommandButton cmdReRaise 
      Caption         =   "Re-Rai&se..."
      Height          =   375
      Left            =   7320
      TabIndex        =   10
      Top             =   7740
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close &Discrepancy..."
      Height          =   375
      Left            =   9960
      TabIndex        =   12
      Top             =   7740
      Width           =   1695
   End
   Begin VB.CommandButton cmdRespond 
      Caption         =   "Res&pond..."
      Height          =   375
      Left            =   8640
      TabIndex        =   11
      Top             =   7740
      Width           =   1215
   End
   Begin VB.TextBox txtMessage 
      Height          =   3195
      Left            =   3900
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   4440
      Width           =   9075
   End
   Begin MSFlexGridLib.MSFlexGrid flxDiscrepancies 
      Height          =   3195
      Left            =   3900
      TabIndex        =   9
      Top             =   4440
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   5636
      _Version        =   393216
      Cols            =   0
      FixedCols       =   0
      AllowUserResizing=   3
   End
   Begin VB.Frame fraQuestion 
      Height          =   4095
      Left            =   60
      TabIndex        =   14
      Top             =   4140
      Width           =   13035
   End
   Begin VB.Frame fraSearch 
      Caption         =   "Search criteria"
      Height          =   4035
      Left            =   60
      TabIndex        =   32
      Top             =   0
      Width           =   13035
   End
   Begin VB.PictureBox picDrag 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H8000000B&
      Height          =   180
      Left            =   60
      MousePointer    =   7  'Size N S
      ScaleHeight     =   180
      ScaleWidth      =   13035
      TabIndex        =   15
      Top             =   3960
      Width           =   13035
   End
End
Attribute VB_Name = "frmViewDiscrepanciesModal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'  Copyright:  InferMed Ltd. 2000. All Rights Reserved
'  File:       frmViewDiscrepancies.frm
'  Author:     Toby Aldridge May 2000
'  Purpose:
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'  Revisions:
'   Mo Morris   26/9/00 Adding Print facilities
'   TA 20/10/2000: Used new GetSQLStringLike for subject label searching
'   TA 30/10/2000: if we come though the view menu we need to restrict on user id when SubjectLabel doesn't exist
'   Mo Morris   24/11/2000  Changes around adding Normal Range and CTC flags to PrintListing and PrintDCF
'                           Changes made to cmdPrintDCF as per the CRC's comments:-
'                           Display of codes removed.
'                           Space added between the line 'On Response', 'Query' and 'Answer'.
'                           Additional space added for the hand written answer.
'               28/11/2000  Changes to PrintListing not made untill 28/11/2000
'               1/12/2000   PrintDCFFooter added for the purpose of adding signing prompt
'                           PrintDCFHeader and PrintListingHeader changed from functions to subs
'   TA 17/01/2001: Changed SQL so that database is only accessed once to populate listview for performance enhancements
'   TA 17/01/2001: Form size and position stored in registry
'   NCJ 20 Feb 01 - Deal with ResponseTimeStamp = 0 in NRCTCText
'
' MACRO 2.2
'   NCJ 25 Sep 01 - Use goArezzo for date parsing in MACRO 2.2
'   TA 27/09/2001: Form made MDI child and did db audit changes
'   TA 03/10/2001: Now filters on the SiteUSer table
'   TA 08/10/2001: OK button removed
'   TA 8/11/01: Clear subject label search box
'
' MACRO 3.0
'   TA 26/11/2001: Started to integrate with new MIMessage Business object
'   TA 05/12/2001: Integrated with new MIMessage busines object - printing still done using the old objects
'   NCJ 11 Mar 02 - Added nRptNo in EFIOpen call in lvwQuestions_DblClick
'                   Set correct caption for cmdClose for Discrepancies
'   TA 8/11/01: Clear subject label search box
'   TA 26/08/2002: Added support for MIMsg statuses
'   TA 26/09/02: Changes for New UI - no title bar, not maximised etc
'   NCJ 14 Oct 02 - Changed Message Status enumeration names
'   NCJ 15-16 Oct 02 - Implemented new statuses for SDVs; consolidation of MIMsg functionality
'   NCJ 18 Oct 02 - Changed some error handlers
'   NCJ 5 Nov 02 - Fixed bugs in Discrepancies due to Timezone additions
'   TA 19/01/2003: Subject locking now done when creating and editing mimessages
'   TA 23/01/2003: moved getting the data from the database into here so refresh will work
'   NCJ 23 Jan 03 - Sorted out display of data in listview (we get "raw" data from GetWinIO.GetMIMessageList)
'   TA 19/02/2003 - Allow user to change mimstatus in monitor mode if they have the eform open
'   TA 14/04/2003 - ensure MIMessage is uptodate when user changes status
'   Mo  9/6/2003    Bug 1836 changes to CreatePrintingSQl and CreateSDVPrintingSQL so that
'                   printing is correct when filtering on visit, eForm and Question.
'   MLM 02/07/03: 3.0 bug list 1885: Pop-up window showing details of Discrepancies and SDVs.
'   DPH 07/11/2003 - LocalNumToStandard date double used in SQL statement in GetNRCTCTextFromResponseHistory
'   DPH 19/01/2004 - Convert date to double when filtering on message created date in CreatePrintingSQL & CreateSDVPrintingSQLFiltering - SR5360
'   MLM 04/07/05: bug 2464: Supply cycle numbers for search and display MIMsgType in title.
'----------------------------------------------------------------------------------------'

Option Explicit

'Private Const m_BACKCOLOR = vbInfoBackground 'yellow
Private Const m_BACKCOLOR = eMACROColour.emcNonWhiteBackGround

Private Const msDEFAULT_PROMPT = "Please enter explanatory text"
'no defualt resond message
Private Const msDEFAULT_RESPOND_MESSAGE = ""
Private Const msDEFAULT_RERAISE_MESSAGE = "Discrepancy re-raised"
Private Const msDEFAULT_CLOSE_MESSAGE = "Discrepancy closed"

Private Const msDEFAULT_DONE_MESSAGE = "Source Data Verification complete"
Private Const msDEFAULT_QUERIED_MESSAGE = "Source Data Verification queried"
Private Const msDEFAULT_CANCELLED_MESSAGE = "Source Data Verification cancelled"
Private Const msDEFAULT_PLANNED_MESSAGE = "Source Data Verification reset to planned"

Private Const msDEFAULT_RESPOND_TITLE = "Discrepancy Response - "
Private Const msDEFAULT_RERAISE_TITLE = "Re-Raise Discrepancy - "
Private Const msDEFAULT_CLOSE_TITLE = "Close Discrepancy - "
Private Const msDEFAULT_DONE_TITLE = "SDV Done - "
Private Const msDEFAULT_PLANNED_TITLE = "SDV Planned - "
Private Const msDEFAULT_QUERIED_TITLE = "SDV Queried - "
Private Const msDEFAULT_CANCELLED_TITLE = "SDV Cancelled - "

'vQI (question identifier) fields:
Private Const m_QI_MessageId = 0
Private Const m_QI_ObjectId = 1
Private Const m_QI_ObjectSource = 2
Private Const m_QI_Site = 3
Private Const m_QI_Text = 4
Private Const m_QI_StudyName = 5
Private Const m_QI_PersonId = 6
Private Const m_QI_Visit = 7
Private Const m_QI_eForm = 8
Private Const m_QI_Question = 9
Private Const m_QI_SubjectLabel = 10
Private Const m_QI_MessageSource = 11
Private Const m_QI_MessageScope = 12

'no min
Private Const mlMIN_HEIGHT = 0 '5000
Private Const mlMIN_WIDTH = 0 '6930

Private Const msglBUTTON_GAP = 90

Private mvQuestions As Variant

Private mnMIMsgType As MIMsgType

Private moMIMessages As clsMIMessages

Private mlIndex As Long

Private msObjectID As String
Private msObjectSource As String

Private moListItem As ListItem

Private mdblProportion As Double
Private mbDrag As Boolean

'Launched from eform or search panel - cannot open efrom from mimsg list if from eForm
Private mnLaunchMode As eMACROWindow

'store the current message object (late bound) - could be note/SDV or discrepancy
Private moMIMsgs As Object 'MIDiscrepancy '-  I declare this as a discrepancy when changing code so that early binding results in a compile error

Private moTimeZone As Timezone

'store an eform passed in if opened from eForm
Private moResponse As Response

'TA 23/01/2003: store the parameters passed in from left hand panel do we can reload after a change
Private mvParams As Variant

'----------------------------------------------------------------------------------------'
Public Sub Display(vParams As Variant, nMIMsgType As MIMsgType, nLaunchMode As eMACROWindow, _
                                Optional oResponse As Response = Nothing)
'----------------------------------------------------------------------------------------'
''VBfnDiscUrl|'+sSt+'|'+sSi+'|'+sVi+'|'+sEf+'|'+sQu+'|'+sSj+'|'+sUs+'|'+sB4+'|'+sTm+'|'+sSs);}"
'                1        2       3       4       5      6       7      8        9      10
''VBfnSDVUrl|'+sSt+'|'+sSi+'|'+sVi+'|'+sEf+'|'+sQu+'|'+sSj+'|'+sUs+'|'+sB4+'|'+sTm+'|'+sSs+'|'+sObj);}"
'                1        2       3       4       5      6       7      8        9      10      11
''VBfnNoteUrl|'+sSt+'|'+sSi+'|'+sVi+'|'+sEf+'|'+sQu+'|'+sSj+'|'+sUs+'|'+sB4+'|'+sTm+'|'+sSs);}"
'                1        2       3       4       5      6       7      8        9      10

'display mimessagelist
'fnDiscUrl(sSt,sSi,sVi,sEf,sQu,sSj,sUs,sB4,sTm,sSs)
'----------------------------------------------------------------------------------------'
Dim vData As Variant
Dim sSubject As String
Dim sPersonId As String


    On Error GoTo Errlabel

    
    If Not frmMenu.SplitScreen Then
        'prompt to close dataentry form and exit if not closed in non split screen mode
        If frmMenu.IsDataEntryFormLoaded(True) Then
'EXIT SUB HERE
            Exit Sub
        End If
    End If

   'TA temp fix   for no study prob
   If vParams(1) = "</select" Then Exit Sub



    sPersonId = ""
    sSubject = vParams(6)



    HourglassOn
    
    'convert mimsgtext to lower case except SDV
    frmHourglass.Display "Loading " & Replace(LCase(GetMIMTypeText(nMIMsgType)), "sdv", "SDV") & " browser", Not frmMenu.SplitScreen

'study id of -1 changed to "" vparams(2)

    ' MLM 01/07/05: explicitly state that we want any repeat of the visit, eform or question..
    Select Case nMIMsgType
    Case MIMsgType.mimtDiscrepancy, MIMsgType.mimtNote
        vData = GetWinIO.GetMIMessageList(nMIMsgType, goUser, IIf(vParams(1) = -1, "", vParams(1)), vParams(2), vParams(3), "" _
                , vParams(4), "", vParams(5), "", vParams(7), sSubject, sPersonId, vParams(10), vParams(9), vParams(8), "")
    
    Case MIMsgType.mimtSDVMark
        ' NCJ 23 Jan 03 - Use sSubject, sPersonId instead of vparams(6)
        vData = GetWinIO.GetMIMessageList(nMIMsgType, goUser, IIf(vParams(1) = -1, "", vParams(1)), vParams(2), vParams(3), "" _
                , vParams(4), "", vParams(5), "", vParams(7), sSubject, sPersonId, vParams(10), vParams(9), vParams(8), vParams(11))
    End Select
    
    Select Case nMIMsgType
    Case MIMsgType.mimtDiscrepancy: OpenWinForm wfDiscepancies
    Case MIMsgType.mimtNote: OpenWinForm wfNotes
    Case MIMsgType.mimtSDVMark: OpenWinForm wfSDV
    End Select
    
    Call SetUpForm(nMIMsgType, nLaunchMode, vData)

    
    frmMenu.Resize
    UnloadfrmHourglass

    
    If IsNull(vData) Then
        DialogInformation "No matching records"
        cmdPrintListing.Enabled = False
        cmdPrintDCF.Enabled = False
    Else
        cmdPrintListing.Enabled = True
        cmdPrintDCF.Enabled = True
    End If

    'store search parameters
    mvParams = vParams


    HourglassOff
    Exit Sub
    
Errlabel:
    Err.Raise Err.Number, , Err.Description & "|frmViewDiscrepancies.Display"
        
End Sub


'----------------------------------------------------------------------------------------'
Private Sub cmdClose_Click()
'----------------------------------------------------------------------------------------'
' Either Close a Discrepancy
' or set an SDV to Done
'----------------------------------------------------------------------------------------'
Dim sMsgText As String
Dim oDisc As MIDiscrepancy
Dim oSDV As MISDV
Dim sResponseValue As String
Dim dblResponseTimeStamp As Double
Dim oLockForMIMsg As clsLockForMIMsg

    On Error GoTo ErrHandler
    
    Set oLockForMIMsg = New clsLockForMIMsg

    If oLockForMIMsg.LockIfNeeded(moMIMsgs.StudyName, moMIMsgs.Site, moMIMsgs.SubjectId, mnMIMsgType, moMIMsgs, moResponse) Then
        'we have a lock or the form open
      If mnMIMsgType = MIMsgType.mimtDiscrepancy Then
          sMsgText = msDEFAULT_CLOSE_MESSAGE
          If frmInputBox.Display(msDEFAULT_CLOSE_TITLE & txtQuestion.Text, msDEFAULT_PROMPT, sMsgText, True, True, True, valOnlySingleQuotes) Then
              Me.Refresh
              Set oDisc = moMIMsgs
              Call GetResponseValueandTimeStamp(moMIMsgs, sResponseValue, dblResponseTimeStamp)
              Call oDisc.CloseDown(sMsgText, goUser.UserName, goUser.UserNameFull, GetMIMsgSource, _
                                  dblResponseTimeStamp, sResponseValue, IMedNow, moTimeZone.TimezoneOffset)
              oDisc.Save
              With oDisc
                  Call UpdateMIMsgStatus(gsADOConnectString, MIMsgType.mimtDiscrepancy, _
                              .StudyName, TrialIdFromName(.StudyName), .Site, .SubjectId, .VisitId, _
                              .VisitCycle, .eFormTaskId, .ResponseTaskId, .ResponseCycle, CurrentSubject)
              End With
    '          cmdRefresh_Click
              
              Call RefreshOrCloseForm
          End If
      Else
          sMsgText = msDEFAULT_DONE_MESSAGE
          If frmInputBox.Display(msDEFAULT_DONE_TITLE & txtQuestion.Text, msDEFAULT_PROMPT, sMsgText, True, True, True, valOnlySingleQuotes) Then
              Me.Refresh
              Set oSDV = moMIMsgs
              Call GetResponseValueandTimeStamp(moMIMsgs, sResponseValue, dblResponseTimeStamp)
              Call oSDV.ChangeStatus(eSDVMIMStatus.ssDone, sMsgText, goUser.UserName, goUser.UserNameFull, GetMIMsgSource, _
                              IMedNow, moTimeZone.TimezoneOffset, dblResponseTimeStamp, sResponseValue)
              oSDV.Save
              With oSDV
                  Call UpdateMIMsgStatus(gsADOConnectString, MIMsgType.mimtSDVMark, _
                              .StudyName, TrialIdFromName(.StudyName), .Site, .SubjectId, .VisitId, _
                              .VisitCycle, .eFormTaskId, .ResponseTaskId, .ResponseCycle, CurrentSubject)
              End With
     '         cmdRefresh_Click
              
              Call RefreshOrCloseForm
          End If
      End If
         
      'unlock if needed
      Call oLockForMIMsg.UnlockIfNeeded
    
    End If
       
Exit Sub
ErrHandler:
    If Err.Number = MIMsgErrors.mimeInvalidForThisStatus Then
        DialogError Err.Description, "Status change unsuccessful"
       'unlock if needed
        Call oLockForMIMsg.UnlockIfNeeded
        RefreshOrCloseForm
    Else

        If MACROErrorHandler(Me.Name, _
                            Err.Number, Err.Description, "cmdClose_Click", Err.Source) = Retry Then
            Resume
        End If
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub cmdCloseForm_Click()
'----------------------------------------------------------------------------------------'
' allow user to close form
'----------------------------------------------------------------------------------------'

    Unload Me

End Sub

'----------------------------------------------------------------------------------------'
Private Sub cmdEdit_Click()
'----------------------------------------------------------------------------------------'
' Allow them to edit the Message text
'----------------------------------------------------------------------------------------'
Dim sMsgText As String
Dim oLockForMIMsg As clsLockForMIMsg

    On Error GoTo ErrHandler
    
   Set oLockForMIMsg = New clsLockForMIMsg
    
    If oLockForMIMsg.LockIfNeeded(moMIMsgs.StudyName, moMIMsgs.Site, moMIMsgs.SubjectId, mnMIMsgType, moMIMsgs, moResponse) Then
        'we have a lock or the form open
    
        sMsgText = moMIMsgs.CurrentMessage.Text
        If frmInputBox.Display("Edit " & GetMIMTypeText(moMIMsgs.MIMessageType) & " - " & txtQuestion.Text, "Edit text", sMsgText, True, True, True, valOnlySingleQuotes) Then
            Me.Refresh
            If sMsgText <> moMIMsgs.CurrentMessage.Text Then
                moMIMsgs.SetText sMsgText, goUser.UserName
                moMIMsgs.Save
                
                RefreshOrCloseForm
            End If
        End If
         
      'unlock if needed
      Call oLockForMIMsg.UnlockIfNeeded
        
    End If
    
Exit Sub
ErrHandler:
    If Err.Number = MIMsgErrors.mimeInvalidForThisStatus Then
        DialogError Err.Description, "Edit unsuccessful"
       'unlock if needed
        Call oLockForMIMsg.UnlockIfNeeded
        RefreshOrCloseForm
    Else

        If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "cmdEdit_Click", Err.Source) = Retry Then
            Resume
        End If
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub cmdPlanned_Click()
'----------------------------------------------------------------------------------------'
' Reset SDV to Planned
'----------------------------------------------------------------------------------------'
Dim sMsgText As String
Dim oSDV As MISDV
Dim sResponseValue As String
Dim dblResponseTimeStamp As Double
Dim oLockForMIMsg As clsLockForMIMsg

    On Error GoTo ErrHandler
    
    Set oLockForMIMsg = New clsLockForMIMsg
    
    If oLockForMIMsg.LockIfNeeded(moMIMsgs.StudyName, moMIMsgs.Site, moMIMsgs.SubjectId, mnMIMsgType, moMIMsgs, moResponse) Then
        'we have a lock or the form open
        
        sMsgText = msDEFAULT_PLANNED_MESSAGE
        If frmInputBox.Display(msDEFAULT_PLANNED_TITLE & txtQuestion.Text, msDEFAULT_PROMPT, sMsgText, True, True, True, valOnlySingleQuotes) Then
            Me.Refresh
            Set oSDV = moMIMsgs
            ' get values for responsevalue and timestamp
            Call GetResponseValueandTimeStamp(moMIMsgs, sResponseValue, dblResponseTimeStamp)
            Call oSDV.ChangeStatus(eSDVMIMStatus.ssPlanned, sMsgText, _
                                goUser.UserName, goUser.UserNameFull, GetMIMsgSource, _
                                IMedNow, moTimeZone.TimezoneOffset, _
                                dblResponseTimeStamp, sResponseValue)
            oSDV.Save
            
            'Update MIMsgStatus
            With oSDV
                Call UpdateMIMsgStatus(gsADOConnectString, MIMsgType.mimtSDVMark, _
                            .StudyName, TrialIdFromName(.StudyName), .Site, .SubjectId, .VisitId, _
                            .VisitCycle, .eFormTaskId, .ResponseTaskId, .ResponseCycle, CurrentSubject)
            End With
            'cmdRefresh_Click
            
            Call RefreshOrCloseForm
        End If
         
      'unlock if needed
      Call oLockForMIMsg.UnlockIfNeeded
    End If

Exit Sub
ErrHandler:
    If Err.Number = MIMsgErrors.mimeInvalidForThisStatus Then
        DialogError Err.Description, "Status change unsuccessful"
           'unlock if needed
        Call oLockForMIMsg.UnlockIfNeeded
        RefreshOrCloseForm
    Else
        If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "cmdPlanned_Click", Err.Source) = Retry Then
            Resume
        End If
    End If

End Sub

'----------------------------------------------------------------------------------------'
Private Sub cmdReRaise_Click()
'----------------------------------------------------------------------------------------'
' This button has a dual function. It will
' either
'   Re-raise a Discrepancy
' or
'   Query an SDV
'----------------------------------------------------------------------------------------'
Dim sMsgText As String
Dim oDisc As MIDiscrepancy
Dim sResponseValue As String
Dim dblResponseTimeStamp As Double
Dim oSDV As MISDV

Dim oLockForMIMsg As clsLockForMIMsg

    On Error GoTo ErrHandler
    
    Set oLockForMIMsg = New clsLockForMIMsg
    
    If oLockForMIMsg.LockIfNeeded(moMIMsgs.StudyName, moMIMsgs.Site, moMIMsgs.SubjectId, mnMIMsgType, moMIMsgs, moResponse) Then
        'we have a lock or the form open
    
        ' Deal with Discrepancy or SDV
        If mnMIMsgType = MIMsgType.mimtDiscrepancy Then
            ' Re-raise discrepancy
            sMsgText = msDEFAULT_RERAISE_MESSAGE
            If frmInputBox.Display(msDEFAULT_RERAISE_TITLE & txtQuestion.Text, msDEFAULT_PROMPT, sMsgText, True, True, True, valOnlySingleQuotes) Then
                Me.Refresh
                Set oDisc = moMIMsgs
                'get values for responsevalue and timestamp
                Call GetResponseValueandTimeStamp(moMIMsgs, sResponseValue, dblResponseTimeStamp)
                Call oDisc.ReRaise(sMsgText, goUser.UserName, goUser.UserNameFull, GetMIMsgSource, _
                                dblResponseTimeStamp, sResponseValue, _
                                IMedNow, moTimeZone.TimezoneOffset)
                oDisc.Save
                
                'Update MIMsgStatus
                With oDisc
                    Call UpdateMIMsgStatus(gsADOConnectString, MIMsgType.mimtDiscrepancy, _
                                .StudyName, TrialIdFromName(.StudyName), .Site, .SubjectId, .VisitId, _
                                .VisitCycle, .eFormTaskId, .ResponseTaskId, .ResponseCycle, CurrentSubject)
                End With
                'cmdRefresh_Click
                
                Call RefreshOrCloseForm
            End If
        Else        ' It's an SDV
            ' Query SDV
            sMsgText = msDEFAULT_QUERIED_MESSAGE
            If frmInputBox.Display(msDEFAULT_QUERIED_TITLE & txtQuestion.Text, msDEFAULT_PROMPT, sMsgText, True, True, True, valOnlySingleQuotes) Then
                Me.Refresh
                Set oSDV = moMIMsgs
                'get values for responsevalue and timestamp
                Call GetResponseValueandTimeStamp(moMIMsgs, sResponseValue, dblResponseTimeStamp)
                Call oSDV.ChangeStatus(eSDVMIMStatus.ssQueried, sMsgText, goUser.UserName, goUser.UserNameFull, _
                                    GetMIMsgSource, IMedNow, moTimeZone.TimezoneOffset, _
                                    dblResponseTimeStamp, sResponseValue)
                oSDV.Save
                
                'Update MIMsgStatus
                With oSDV
                    Call UpdateMIMsgStatus(gsADOConnectString, MIMsgType.mimtSDVMark, _
                                .StudyName, TrialIdFromName(.StudyName), .Site, .SubjectId, .VisitId, _
                                .VisitCycle, .eFormTaskId, .ResponseTaskId, .ResponseCycle, CurrentSubject)
                End With
                'cmdRefresh_Click
                
                Call RefreshOrCloseForm
            End If
        End If
         
      'unlock if needed
      Call oLockForMIMsg.UnlockIfNeeded
    End If
        
Exit Sub

ErrHandler:
    If Err.Number = MIMsgErrors.mimeInvalidForThisStatus Then
        DialogError Err.Description, "Status change unsuccessful"
       'unlock if needed
        Call oLockForMIMsg.UnlockIfNeeded
        RefreshOrCloseForm
    Else
    
        If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "cmdReRaise_Click", Err.Source) = Retry Then
            Resume
        End If
    End If
End Sub

'----------------------------------------------------------------------------------------'
Private Sub cmdRespond_Click()
'----------------------------------------------------------------------------------------'
' This button has a dual function. It will
' either
'   Respond to a Discrepancy
' or
'   Cancel an SDV
'----------------------------------------------------------------------------------------'
Dim sMsgText As String
Dim oDisc As MIDiscrepancy
Dim oSDV As MISDV
Dim sResponseValue As String
Dim dblResponseTimeStamp As Double

Dim oLockForMIMsg As clsLockForMIMsg

    On Error GoTo ErrHandler
    
    Set oLockForMIMsg = New clsLockForMIMsg
    
    If oLockForMIMsg.LockIfNeeded(moMIMsgs.StudyName, moMIMsgs.Site, moMIMsgs.SubjectId, mnMIMsgType, moMIMsgs, moResponse) Then
        'we have a lock or the form open
            
        ' Deal with Discrepancy or SDV
        If mnMIMsgType = MIMsgType.mimtDiscrepancy Then
            sMsgText = msDEFAULT_RESPOND_MESSAGE
            If frmInputBox.Display(msDEFAULT_RESPOND_TITLE & txtQuestion.Text, msDEFAULT_PROMPT, sMsgText, True, True, True, valOnlySingleQuotes) Then
                Me.Refresh
                Set oDisc = moMIMsgs
                'get values for responsevalue and timestamp
                Call GetResponseValueandTimeStamp(moMIMsgs, sResponseValue, dblResponseTimeStamp)
                Call oDisc.Respond(sMsgText, goUser.UserName, goUser.UserNameFull, GetMIMsgSource, _
                                dblResponseTimeStamp, sResponseValue, _
                                IMedNow, moTimeZone.TimezoneOffset)
                oDisc.Save
                'Update MIMsgStatus
                With oDisc
                    Call UpdateMIMsgStatus(gsADOConnectString, MIMsgType.mimtDiscrepancy, _
                                .StudyName, TrialIdFromName(.StudyName), .Site, .SubjectId, .VisitId, _
                                .VisitCycle, .eFormTaskId, .ResponseTaskId, .ResponseCycle, CurrentSubject)
                End With
                'cmdRefresh_Click
                
                Call RefreshOrCloseForm
    
            End If
        Else        ' It's an SDV
            ' Cancel SDV
            sMsgText = msDEFAULT_CANCELLED_MESSAGE
            If frmInputBox.Display(msDEFAULT_CANCELLED_TITLE & txtQuestion.Text, msDEFAULT_PROMPT, sMsgText, True, True, True, valOnlySingleQuotes) Then
                Me.Refresh
                Set oSDV = moMIMsgs
                'get values for responsevalue and timestamp
                Call GetResponseValueandTimeStamp(moMIMsgs, sResponseValue, dblResponseTimeStamp)
                Call oSDV.ChangeStatus(eSDVMIMStatus.ssCancelled, sMsgText, goUser.UserName, goUser.UserNameFull, _
                                    GetMIMsgSource, IMedNow, moTimeZone.TimezoneOffset, _
                                    dblResponseTimeStamp, sResponseValue)
                oSDV.Save
                
                'Update MIMsgStatus
                With oSDV
                    Call UpdateMIMsgStatus(gsADOConnectString, MIMsgType.mimtSDVMark, _
                                .StudyName, TrialIdFromName(.StudyName), .Site, .SubjectId, .VisitId, _
                                .VisitCycle, .eFormTaskId, .ResponseTaskId, .ResponseCycle, CurrentSubject)
                End With
    
                
                Call RefreshOrCloseForm
            End If
        End If
         
      'unlock if needed
      Call oLockForMIMsg.UnlockIfNeeded
    End If
    
Exit Sub
ErrHandler:
    If Err.Number = MIMsgErrors.mimeInvalidForThisStatus Then
        DialogError Err.Description, "Status change unsuccessful"
       'unlock if needed
        Call oLockForMIMsg.UnlockIfNeeded
        RefreshOrCloseForm
    Else
    
        If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "cmdRespond_Click", Err.Source) = Retry Then
            Resume
        End If
    End If

End Sub

'----------------------------------------------------------------------------------------
Private Sub cmdShowDetails_Click()
'----------------------------------------------------------------------------------------
'MLM 01/07/03: Created. Display a modal window to show the history of the selected MIMessage
'   in a bigger area.
'----------------------------------------------------------------------------------------

Dim ofrmMIMsg As frmWebNonMDI

    Set ofrmMIMsg = New frmWebNonMDI
       
    With ofrmMIMsg
        .Width = 7000
        .Height = 4000
        .Display wdtHTML, MIMessageHistoryHTML(moMIMsgs), "auto", True, GetMIMTypeText(moMIMsgs.MIMessageType) & " - " & txtQuestion.Text
    End With

    Set ofrmMIMsg = Nothing

End Sub

'----------------------------------------------------------------------------------------'
Private Sub Form_Load()
'----------------------------------------------------------------------------------------'
'One time only code
'----------------------------------------------------------------------------------------'
Dim conControl As Control

    'st icon
    Me.Icon = frmMenu.Icon
    
    'set correct colours for controls

    picScope.BackColor = m_BACKCOLOR

    
    fraQuestion.BackColor = eMACROColour.emcBackground
    fraSearch.BackColor = eMACROColour.emcBackground
    
    'set the background colour for all checkboxes, option buttons, textboxes and labels
    For Each conControl In Me.Controls
         If Left(conControl.Name, 3) = "chk" Or Left(conControl.Name, 3) = "opt" Or Left(conControl.Name, 3) = "txt" Or Left(conControl.Name, 3) = "lbl" Then
            conControl.BackColor = m_BACKCOLOR
        End If
    Next
    'except the date box, the subject label search box and the message text box

    txtMessage.BackColor = eMACROColour.emcBackground

    
    Me.BackColor = eMACROColour.emcBackground
    
    'where the drag bar appears in the form
    mdblProportion = 0.5
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mnLaunchMode = SubjectMIMEssage Then
        'store window dimensions
        Call SaveFormDimensions(Me)
    End If
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Form_Unload(Cancel As Integer)
'----------------------------------------------------------------------------------------'
Dim ofrm As Form

    mnLaunchMode = None
    Set moMIMsgs = Nothing
    Set moTimeZone = Nothing
    Set moResponse = Nothing
    
    'inform borders that i am closing
    Select Case mnMIMsgType
    Case mimtDiscrepancy: CloseWinForm wfDiscepancies
    Case mimtNote: CloseWinForm wfNotes
    Case mimtSDVMark: CloseWinForm wfSDV
    End Select
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub lblPriority_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'----------------------------------------------------------------------------------------'
Dim nPriority As Long
Dim oDisc As MIDiscrepancy

Dim oLockForMIMsg As clsLockForMIMsg

    On Error GoTo ErrHandler
    
    Set oLockForMIMsg = New clsLockForMIMsg
    
    If oLockForMIMsg.LockIfNeeded(moMIMsgs.StudyName, moMIMsgs.Site, moMIMsgs.SubjectId, mnMIMsgType, moMIMsgs, moResponse) Then
        'we have a lock or the form open
    
        Set oDisc = moMIMsgs
        
        If Not oDisc Is Nothing Then
            If lblPriority.TooltipText = "Right click to change priority" Then
                If Button = vbRightButton Then
                    nPriority = frmMenu.ShowPopUp("1|2|3|4|5|6|7|8|9|10")
                    If nPriority <> -1 Then
                        oDisc.SetPriority CInt(nPriority), GetMIMsgSource
                        oDisc.Save
                        lblPriority.Caption = nPriority
                        'cmdRefresh_Click
                    End If
                End If
            End If
        End If
         
      'unlock if needed
      Call oLockForMIMsg.UnlockIfNeeded
        
    End If
Exit Sub
ErrHandler:
    If Err.Number = MIMsgErrors.mimeInvalidForThisStatus Then
        DialogError Err.Description, "Priority change unsuccessful"
       'unlock if needed
        Call oLockForMIMsg.UnlockIfNeeded
        RefreshOrCloseForm
    Else
    
        If MACROErrorHandler("frmViewDiscrepancies", _
                            Err.Number, Err.Description, "lblPriority_MouseUp", Err.Source) = Retry Then
            Resume
        End If
    End If
End Sub

'----------------------------------------------------------------------------------------'
Private Function PopulateListView(vData As Variant) As Long
'----------------------------------------------------------------------------------------'
' fill the listview with MIMessage details according to sSQL string
'returns number of rows in the listview
' NCJ 14 Oct 02 - Added Queried and Cancelled for SDVs
' NCJ 23 Jan 03 - Sorted out massaging of "raw" data for display
'----------------------------------------------------------------------------------------'
Dim rs As ADODB.Recordset
Dim seForm As String
Dim sVisit As String
Dim sQuestion As String
Dim sMsgSummary As String
Dim lStudy As Long
Dim i As Long
Dim lSelect As Long
Dim sFilter As String
Dim sIn As String
Dim vQI As Variant

Dim oMIMList As New MIDataLists
Dim sStudyName As String
Dim sSite As String
Dim lVisitId As Long
Dim lSubjectId As Long
Dim sSubjectLabel As String
Dim sUser As String
Dim vStatus As Variant
Dim dblDate As Double
Dim bBefore As Boolean
Dim colStatus As Collection
Dim lEFormId As Long
Dim lQuId As Long
Dim sOCIdHeader As String

    On Error GoTo Errlabel

    HourglassOn

    'TA 18/11/2004: flag to show OC Ids isssue 2448
    'default to true
    If (LCase(GetMACROSetting(MACRO_SETTING_USE_OC_ID, "true")) = "true") Then
        sOCIdHeader = "OC Id"
    Else
        sOCIdHeader = "Id"
    End If
    
    Set rs = New ADODB.Recordset
    
    ReDim mvQuestions(0) As String
    
    With rs
        Select Case mnMIMsgType
        Case MIMsgType.mimtDiscrepancy, MIMsgType.mimtSDVMark
            If mnMIMsgType = MIMsgType.mimtDiscrepancy Then
                .Fields.Append "Priority", adVarChar, 2, adFldIsNullable
            End If
            .Fields.Append "Date", adVarChar, 50, adFldIsNullable
            If mnMIMsgType = MIMsgType.mimtSDVMark Then     ' NCJ 15 Oct 02
                .Fields.Append "Scope", adVarChar, 255, adFldIsNullable
            End If
            .Fields.Append "Status", adVarChar, 50, adFldIsNullable
            .Fields.Append "Subject", adVarChar, 50, adFldIsNullable  'TA 07/08/2000 SR3763
            ' NCJ 15 Oct 02 - If SDV, show Visit name
            If mnMIMsgType = MIMsgType.mimtSDVMark Then
                .Fields.Append "Visit", adVarChar, 255, adFldIsNullable
            End If
            .Fields.Append "eForm", adVarChar, 255, adFldIsNullable
            .Fields.Append "Question", adVarChar, 255, adFldIsNullable
            .Fields.Append "User Name", adVarChar, 50, adFldIsNullable
            'TA 12/06/2000 SR 3582: show OC Discrepancy Id
            If mnMIMsgType = MIMsgType.mimtDiscrepancy Then
'TA 18/11/2004: header according to show OC ids flag, isssue 2448
                .Fields.Append sOCIdHeader, adVarChar, 50, adFldIsNullable
            End If
            .Fields.Append "Text", adVarChar, 2000, adFldIsNullable
            .Open , , adOpenKeyset, adLockOptimistic
        Case MIMsgType.mimtNote
            .Fields.Append "Time Stamp", adVarChar, 50, adFldIsNullable
            .Fields.Append "Subject", adVarChar, 50, adFldIsNullable    'TA 07/08/2000 SR3763
            .Fields.Append "eForm", adVarChar, 255, adFldIsNullable
            .Fields.Append "Question", adVarChar, 255, adFldIsNullable
            .Fields.Append "User Name", adVarChar, 50, adFldIsNullable
            .Fields.Append "Status", adVarChar, 10, adFldIsNullable
            .Fields.Append "Text", adVarChar, 2000, adFldIsNullable
            .Open , , adOpenKeyset, adLockOptimistic
        End Select
    
        If Not IsNull(vData) Then
        
            For i = 0 To UBound(vData, 2)
                ' NCJ 23 Jan 03 - Initialise Visit, eForm, Question for each row
                sVisit = ""
                seForm = ""
                sQuestion = ""
                lStudy = vData(mmcStudyid, i)
                If vData(mmcVisitId, i) <> 0 Then
                    'if we have a visit
                    sVisit = goUser.DataLists.GetStudyItemName(soVisit, lStudy, CLng(vData(mmcVisitId, i))) _
                                & CycleNumberText(vData(MIMsgCol.mmcVisitCycle, i))
                End If
                If vData(mmcEFormId, i) <> 0 Then
                    'if we have an eform
                    seForm = goUser.DataLists.GetStudyItemName(soeform, lStudy, CLng(vData(mmcEFormId, i))) _
                                & CycleNumberText(vData(MIMsgCol.mmcEFormCycle, i))
                End If
                If vData(mmcQuestionId, i) <> 0 Then
                    'if we have a question
                    sQuestion = goUser.DataLists.GetStudyItemName(soQuestion, lStudy, CLng(vData(mmcQuestionId, i))) _
                                & CycleNumberText(vData(MIMsgCol.mmcResponseCycle, i))
                End If
    
                sMsgSummary = vData(MIMsgCol.mmcId, i) & "|" _
                            & vData(MIMsgCol.mmcObjectId, i) & "|" _
                            & vData(MIMsgCol.mmcObjectSource, i) & "|" _
                            & vData(MIMsgCol.mmcSite, i) & "|" _
                            & vData(MIMsgCol.mmcText, i) & "|" _
                            & vData(MIMsgCol.mmcStudyName, i) & "|" _
                            & vData(MIMsgCol.mmcSubjectId, i) & "|" _
                            & sVisit & "|" _
                            & seForm & "|" & sQuestion & "|" _
                            & RemoveNull(vData(MIMsgCol.mmcSubjectLabel, i)) _
                            & "|" & vData(MIMsgCol.mmcSource, i) _
                            & "|" & vData(MIMsgCol.mmcScope, i)     ' NCJ 15 Oct 02 Added Scope

                ReDim Preserve mvQuestions(UBound(mvQuestions) + 1)
                mvQuestions(UBound(mvQuestions)) = sMsgSummary
                
                Select Case mnMIMsgType
                Case MIMsgType.mimtDiscrepancy
                        .AddNew
                        .Fields("Priority") = vData(MIMsgCol.mmcPrioirty, i)
                        .Fields("Date") = Format(vData(MIMsgCol.mmcCreated, i), "yyyy/mm/dd")
                        .Fields("Status") = MACROMIMsgBS30.GetStatusText(MIMsgType.mimtDiscrepancy, (vData(MIMsgCol.mmcStatus, i)))
                        '.Fields("Subject") = RemoveNull(vData(MIMsgCol.mmcSubjectLabel, i))
                        .Fields("Subject") = vData(MIMsgCol.mmcStudyName, i) _
                                                & "/" & vData(MIMsgCol.mmcSite, i) _
                                                & "/" & RtnSubjectText(vData(MIMsgCol.mmcSubjectId, i), vData(MIMsgCol.mmcSubjectLabel, i))
                        .Fields("eForm") = seForm
                        .Fields("Question") = sQuestion
                        .Fields("User Name") = vData(MIMsgCol.mmcUserNameFull, i)
                        'TA 12/06/2000 SR 3582: show OC Discrepancy Id
                        If vData(MIMsgCol.mmcExternalId, i) <> 0 Then
                            'TA 18/11/2004: header according to show OC ids flag, isssue 2448
                            .Fields(sOCIdHeader) = vData(MIMsgCol.mmcExternalId, i)
                        End If
                        .Fields("Text") = vData(MIMsgCol.mmcText, i)
                Case MIMsgType.mimtSDVMark
                        .AddNew
                        .Fields("Date") = Format(vData(MIMsgCol.mmcCreated, i), "yyyy/mm/dd")
                        .Fields("Scope") = MACROMIMsgBS30.GetScopeText(vData(MIMsgCol.mmcScope, i))
                        .Fields("Status") = MACROMIMsgBS30.GetStatusText(MIMsgType.mimtSDVMark, (vData(MIMsgCol.mmcStatus, i)))
                        '.Fields("Subject") = RemoveNull(vData(MIMsgCol.mmcSubjectLabel, i))
                        .Fields("Subject") = vData(MIMsgCol.mmcStudyName, i) _
                                                & "/" & vData(MIMsgCol.mmcSite, i) _
                                                & "/" & RtnSubjectText(vData(MIMsgCol.mmcSubjectId, i), vData(MIMsgCol.mmcSubjectLabel, i))
                        .Fields("Visit") = sVisit
                        .Fields("eForm") = seForm
                        .Fields("Question") = sQuestion
                        .Fields("User Name") = vData(MIMsgCol.mmcUserNameFull, i)
                        .Fields("Text") = vData(MIMsgCol.mmcText, i)
                Case MIMsgType.mimtNote
                        .AddNew
                        '.Fields("Subject") = RemoveNull(vData(MIMsgCol.mmcSubjectLabel, i))
                        .Fields("Subject") = vData(MIMsgCol.mmcStudyName, i) _
                                                & "/" & vData(MIMsgCol.mmcSite, i) _
                                                & "/" & RtnSubjectText(vData(MIMsgCol.mmcSubjectId, i), vData(MIMsgCol.mmcSubjectLabel, i))
                        .Fields("eForm") = seForm
                        .Fields("Question") = sQuestion
                        .Fields("User Name") = vData(MIMsgCol.mmcUserNameFull, i)
                        .Fields("Text") = vData(MIMsgCol.mmcText, i)
                        .Fields("Time Stamp") = Format(vData(MIMsgCol.mmcCreated, i), "yyyy/mm/dd hh:mm:ss")
                        .Fields("Status") = MACROMIMsgBS30.GetStatusText(MIMsgType.mimtNote, (vData(MIMsgCol.mmcStatus, i)))
                End Select
    
            Next
        End If
        
    End With
    lvwQuestions.Visible = False
    Call RecordSet_toListView(lvwQuestions, rs)
    lvwQuestions.Visible = True
    
    'could display the number of records returned if I could find somewhere to put it
    'lblRecords.Caption = nRecords & " record(s) found"

    rs.Close
    Set rs = Nothing
    
    mlIndex = 0
    
    'select old item by object source and id
    If lvwQuestions.ListItems.Count > 0 Then
        lSelect = 1
         For i = 1 To UBound(mvQuestions)
            vQI = Split(mvQuestions(i), "|")
             If (vQI(m_QI_ObjectId) = msObjectID) And (vQI(m_QI_ObjectSource) = msObjectSource) Then
                 lSelect = i
                 Exit For
             End If
         Next
        For i = 1 To lvwQuestions.ListItems.Count
            If lvwQuestions.ListItems(i).Tag = lSelect Then
                lvwQuestions.SelectedItem = lvwQuestions.ListItems(i)
                Call QuestionSelect(Val(lvwQuestions.ListItems(i).Tag))
                On Error Resume Next
                lvwQuestions.SetFocus
                Err.Clear
                On Error GoTo Errlabel
                Exit For
            End If
        Next
    Else
        txtMessage.Text = ""
        flxDiscrepancies.Cols = 0
        flxDiscrepancies.Rows = 0
        txtStudy.Text = ""
        txtSite.Text = ""
        'TA 20/10/2000: ensure subject label blanked
        txtSubjectLabel.Text = ""
        txtPerson.Text = ""
        txtVisit.Text = ""
        txteForm.Text = ""
        txtQuestion.Text = ""
        txtStatus.Text = ""
        lblPriority.Caption = ""
        Set moMIMsgs = Nothing
        View
    End If
    
    HourglassOff
    
    PopulateListView = lvwQuestions.ListItems.Count
       
Exit Function
Errlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "frmViewDiscrepancies.PopulateListView"
            
End Function

'----------------------------------------------------------------------------------------'
Private Function CurrentSubject() As StudySubject
'----------------------------------------------------------------------------------------'
' Get the currently loaded subject, if any
'----------------------------------------------------------------------------------------'
    
    If Not goStudyDef Is Nothing Then
        Set CurrentSubject = goStudyDef.Subject
    Else
        Set CurrentSubject = Nothing
    End If
   
End Function

'----------------------------------------------------------------------------------------'
Public Sub DisplayModal(nMIMsgType As MIMsgType, nLaunchMode As eMACROWindow, vData As Variant, _
                                Optional oResponse As Response = Nothing)
'----------------------------------------------------------------------------------------'
'modal version of display
'----------------------------------------------------------------------------------------'

    'disable the 2 printing command buttons when in Modal mode
    '(because the variant array of parameters for the printout SQL is not available)
    cmdPrintListing.Visible = False
    cmdPrintDCF.Visible = False
    
    Call SetUpForm(nMIMsgType, nLaunchMode, vData, oResponse)
    
End Sub


'----------------------------------------------------------------------------------------'
Private Sub SetUpForm(nMIMsgType As MIMsgType, nLaunchMode As eMACROWindow, vData As Variant, _
                                Optional oResponse As Response = Nothing)
'----------------------------------------------------------------------------------------'
' called externally to display form
'----------------------------------------------------------------------------------------'
Dim conControl As Control
Dim bChanged As Boolean ' are we refreshing the display
Dim ofrm As Form

    On Error GoTo ErrHandler
        
    Load Me
    
    Set moResponse = oResponse
   
    Set moTimeZone = New Timezone


    mnLaunchMode = nLaunchMode
    mnMIMsgType = nMIMsgType

    cmdPrintListing.Enabled = False
    
    'TA 13/11/01: ensure these buttons start visible (code will hide them later if needed)
    cmdReRaise.Visible = True
    cmdClose.Visible = True
    cmdRespond.Visible = True
    cmdPlanned.Visible = True
    cmdShowDetails.Visible = True
    
    'clear everything
    lvwQuestions.ListItems.Clear
    lvwQuestions.ColumnHeaders.Clear
    flxDiscrepancies.Clear
    txtMessage.Text = ""
    flxDiscrepancies.Cols = 0
    flxDiscrepancies.Rows = 0
    txtStudy.Text = ""
    txtSite.Text = ""
    'TA 20/10/2000: ensure subject label blanked
    txtSubjectLabel.Text = ""
    txtPerson.Text = ""
    txtVisit.Text = ""
    txteForm.Text = ""
    txtQuestion.Text = ""
    txtStatus.Text = ""
    lblPriority.Caption = ""
    Set moMIMsgs = Nothing
     
    'hide everything
    
    txtMessage.Visible = False
    flxDiscrepancies.Visible = False
     
    Select Case mnMIMsgType
    Case MIMsgType.mimtDiscrepancy

        If nLaunchMode <> SubjectMIMEssage Then
            cmdPrintDCF.Visible = True
            cmdPrintDCF.Enabled = False
        End If
        cmdPlanned.Visible = False

        
        'resize close button
        cmdClose.Width = 1695
        ' Rename buttons (in case they were used by SDVs)
        cmdClose.Caption = "Close &Discrepancy..."
        cmdRespond.Caption = "Res&pond..."
        cmdReRaise.Caption = "Re-rai&se..."
       
        'close button may have been resized - check all are in place
        Call PlaceButtons
        DoEvents
        
        flxDiscrepancies.Visible = True

        
    Case MIMsgType.mimtSDVMark

        cmdClose.Width = 1215
        ' Rename buttons (in case they were used by Discrepancies)
        cmdClose.Caption = "&Done..."
        cmdRespond.Caption = "&Cancelled..."
        cmdReRaise.Caption = "&Queried..."
        
        'close button may have been resized - check all are in place
        Call PlaceButtons
        DoEvents
        
        lblPriority1.Visible = False
        lblPriority.Visible = False
        'Mo Morris 26/9/00
        cmdPrintDCF.Visible = False
        
        'turn on appropriate controls
        flxDiscrepancies.Visible = True

        
    Case MIMsgType.mimtNote

        'msUsercode = gUser.UserName
      
        cmdReRaise.Visible = False
        cmdClose.Visible = False
        cmdRespond.Visible = False
        cmdPlanned.Visible = False
        cmdShowDetails.Visible = False

        lblStatus.Visible = False
        txtStatus.Visible = False
        lblPriority1.Visible = False
        lblPriority.Visible = False
        'Mo Morris 26/9/00
        cmdPrintDCF.Visible = False
                    
        'turn on appropriate controls
        txtMessage.Visible = True

    End Select

    Call View

    'ensure tasklist uptodate
    Call frmMenu.UpdateDiscCount
    Call PopulateListView(vData)
    
    Me.WindowState = vbNormal
    
    If nLaunchMode = MonitorMIMessage Then
        Me.Show vbModeless
        Me.ZOrder
    Else
        'MLM 04/07/05: Just show the MIMsgType, if we weren't launched from an eForm
        If moResponse Is Nothing Then
            Me.Caption = GetMIMTypeText(mnMIMsgType, True)
        Else
            Me.Caption = GetMIMTypeText(mnMIMsgType, True) & " for question " & moResponse.Element.Name
        End If
        Me.BorderStyle = vbSizable
        'set size, position and window state
        Call SetFormDimensions(Me)
        Me.Show vbModal
    End If
    
Exit Sub
ErrHandler:
    If MACROErrorHandler("frmViewDiscrepancies", _
                        Err.Number, Err.Description, "Display", Err.Source) = Retry Then
        Resume
    End If

End Sub

'----------------------------------------------------------------------------------------'
Private Sub PlaceButtons()
'----------------------------------------------------------------------------------------'
' Place the buttons in their correct positions
'----------------------------------------------------------------------------------------'

    cmdEdit.Left = flxDiscrepancies.Left + flxDiscrepancies.Width - cmdEdit.Width
    cmdClose.Left = cmdEdit.Left - (cmdClose.Width + msglBUTTON_GAP)
    cmdRespond.Left = cmdEdit.Left - (cmdClose.Width + cmdEdit.Width + 2 * msglBUTTON_GAP)
    cmdReRaise.Left = cmdEdit.Left - (cmdClose.Width + cmdEdit.Width + cmdRespond.Width + 3 * msglBUTTON_GAP)
    cmdPlanned.Left = cmdEdit.Left - (cmdClose.Width + cmdEdit.Width + cmdRespond.Width + cmdReRaise.Width + 4 * msglBUTTON_GAP)
    'MLM 30/06/03:
    cmdShowDetails.Left = cmdEdit.Left - (cmdClose.Width + cmdEdit.Width + cmdRespond.Width + cmdReRaise.Width + 4 * msglBUTTON_GAP + _
        IIf(cmdPlanned.Visible, cmdPlanned.Width + msglBUTTON_GAP, 0))

End Sub

'----------------------------------------------------------------------------------------'
Private Sub Form_Resize()
'----------------------------------------------------------------------------------------'

    On Error Resume Next
    
      If Me.Height >= mlMIN_HEIGHT Then
        
        picDrag.Top = mdblProportion * (Me.Height - cmdPrintListing.Height - 240)
        
        fraSearch.Height = picDrag.Top - 60
        lvwQuestions.Height = fraSearch.Height - 360
        

        
        fraQuestion.Top = picDrag.Top + 60
        fraQuestion.Height = Me.ScaleHeight - fraQuestion.Top - cmdPrintListing.Height - 120
        flxDiscrepancies.Top = fraQuestion.Top + 240
        flxDiscrepancies.Height = fraQuestion.Height - cmdPrintListing.Height - 420
        picScope.Top = flxDiscrepancies.Top
       
        cmdPrintListing.Top = fraQuestion.Top + fraQuestion.Height + 60
        'Mo Morris 26/9/000
        cmdPrintDCF.Top = cmdPrintListing.Top
        cmdCloseForm.Top = cmdPrintListing.Top
        
        cmdEdit.Top = flxDiscrepancies.Top + flxDiscrepancies.Height + 60
        cmdReRaise.Top = cmdEdit.Top
        cmdClose.Top = cmdEdit.Top
        cmdRespond.Top = cmdEdit.Top
        cmdPlanned.Top = cmdEdit.Top
        'MLM 30/06/03:
        cmdShowDetails.Top = cmdEdit.Top
        
        'if message or note
        txtMessage.Top = fraQuestion.Top + 240
        txtMessage.Height = fraQuestion.Height - cmdPrintListing.Height - 420
        
        picScope.Height = flxDiscrepancies.Height

        
    End If
        
    If Me.Width >= mlMIN_WIDTH Then
        fraSearch.Width = Me.ScaleWidth - 120
        lvwQuestions.Width = fraSearch.Width - 300
        lvwQuestions.Width = fraSearch.Width - 120
        
        fraQuestion.Width = fraSearch.Width
        flxDiscrepancies.Width = fraQuestion.Width - picScope.Width - 360
        
        picDrag.Width = fraQuestion.Width

        Call PlaceButtons
        
        'if message or note
        txtMessage.Width = fraQuestion.Width - picScope.Width - 360
        
        cmdCloseForm.Left = Me.Width - cmdCloseForm.Width - 300
    
    End If
    
    picDrag.BackColor = eMACROColour.emcTitlebar
    picDrag.BorderStyle = vbSolid
Exit Sub

ErrHandler:
    
End Sub


'----------------------------------------------------------------------------------------'
Private Sub lvwQuestions_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'----------------------------------------------------------------------------------------'
' sort listview according to column click
'----------------------------------------------------------------------------------------'
    
    lvw_Sort lvwQuestions, ColumnHeader
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub QuestionSelect(lIndex As Long)
'----------------------------------------------------------------------------------------'
'for SDV and Discrepancy fill the the grid with Question SDV or Discrepancy details
'for Message and Note fill the textbox with Question Message/Note text
'----------------------------------------------------------------------------------------'
Dim vQI As Variant
Dim sSQL As String
Dim rsDiscrepancies As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim oMIMessage As clsMIMessage
Dim oMIMsg As MIMsg
Dim oDisc As MIDiscrepancy
Dim oSDV As MISDV
Dim oNote As MINote

    On Error GoTo ErrHandler

    vQI = Split(mvQuestions(lIndex), "|")

    msObjectID = vQI(m_QI_ObjectId)
    msObjectSource = vQI(m_QI_ObjectSource)
    
'If True Then 'DialogQuestion("new?") = vbYes Then

    Select Case mnMIMsgType
    Case MIMsgType.mimtDiscrepancy
        Set oDisc = New MIDiscrepancy
        oDisc.Load gsADOConnectString, CLng(msObjectID), CLng(msObjectSource), (vQI(m_QI_Site))
        Set moMIMsgs = oDisc
    Case MIMsgType.mimtSDVMark
        Set oSDV = New MISDV
        oSDV.Load gsADOConnectString, CLng(msObjectID), CLng(msObjectSource), (vQI(m_QI_Site))
        Set moMIMsgs = oSDV
    Case MIMsgType.mimtNote
        Set oNote = New MINote
        oNote.Load gsADOConnectString, CLng(vQI(m_QI_MessageId)), CLng(vQI(m_QI_MessageSource)), (vQI(m_QI_Site))
        Set moMIMsgs = oNote
    End Select
    
    Select Case mnMIMsgType
    Case MIMsgType.mimtDiscrepancy, MIMsgType.mimtSDVMark
    
        Set rs = New ADODB.Recordset

        With rs
            .Fields.Append "Created", adVarChar, 50, adFldIsNullable
            .Fields.Append "Status", adVarChar, 50, adFldIsNullable
            .Fields.Append "Text", adVarChar, 2000, adFldIsNullable
            .Fields.Append "Value", adVarChar, 255, adFldIsNullable
            .Fields.Append "User Name", adVarChar, 50, adFldIsNullable
            .Open , , adOpenKeyset, adLockOptimistic

            For Each oMIMsg In moMIMsgs.Messages
                .AddNew
                .Fields("Created") = Format(oMIMsg.TimeCreated, "yyyy/mm/dd hh:mm:ss")
                .Fields("Status") = oMIMsg.StatusText
                .Fields("Text") = oMIMsg.Text
                .Fields("Value") = oMIMsg.ResponseValue
                .Fields("User Name") = oMIMsg.UserNameFull
            Next
        End With

        rs.Sort = "Created ASC"
        RecordSet_toGrid flxDiscrepancies, rs, , "||30|30|", True
        rs.Close
        Set rs = Nothing
    Case Else
        txtMessage.Text = oNote.CurrentMessage.Text
    End Select
    
    If mnMIMsgType = MIMsgType.mimtDiscrepancy Then

       txtStatus.Text = moMIMsgs.CurrentMessage.StatusText
       lblPriority.Caption = moMIMsgs.CurrentMessage.Priority
    Else
       txtStatus.Text = moMIMsgs.CurrentMessage.StatusText
    End If

    txtStudy.Text = vQI(m_QI_StudyName)
    txtSite.Text = vQI(m_QI_Site)
    txtPerson.Text = vQI(m_QI_PersonId)
    txtSubjectLabel.Text = vQI(m_QI_SubjectLabel)   'TA 07/08/2000 SR3763
    txtVisit.Text = vQI(m_QI_Visit)
    txteForm.Text = vQI(m_QI_eForm)
    txtQuestion.Text = vQI(m_QI_Question)
    
    Call View
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "QuestionSelect")
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
Private Sub lvwQuestions_DblClick()
'----------------------------------------------------------------------------------------'
' Display the eForm for the selected response
' NCJ 11 Mar 02 - Added ResponseCycle in EFIOpen call
' NCJ 15 Oct 02 - Only show eForm for Question or eForm MIMsg
'----------------------------------------------------------------------------------------'
Dim lIndex As Long
Dim vMsgInfo As Variant
Dim enScope As MIMsgScope

    On Error GoTo ErrHandler
        
    If mnLaunchMode = MonitorMIMessage Then
        'only allow eform opening if they have opened this form from the search panel
                
        If Not (lvwQuestions.SelectedItem Is Nothing) Then
            HourglassOn
            lIndex = lvwQuestions.SelectedItem.Tag
            vMsgInfo = Split(mvQuestions(lIndex), "|")
            
            enScope = vMsgInfo(m_QI_MessageScope)
            ' We can only display an eForm for a Question or eForm MIMsg
            
            Select Case enScope
            Case MIMsgScope.mimscQuestion
                With moMIMsgs
                    ' NCJ 11 Mar 02 - Added .ResponseCycle to EFIOpen
                    Call frmMenu.EFIOpen(TrialIdFromName(.StudyName), .Site, .SubjectId, _
                            .eFormTaskId, .ResponseTaskId, .ResponseCycle, "", True)
                End With
                
            Case MIMsgScope.mimscEForm
                With moMIMsgs
                    Call frmMenu.EFIOpen(TrialIdFromName(.StudyName), .Site, .SubjectId, _
                            .eFormTaskId, glMINUS_ONE, 1, "", True)
                End With
                
            Case Else
                ' We don't do anything
            End Select
            HourglassOff
        End If
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "lvwQuestions_DblClick()")
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
Private Sub lvwQuestions_ItemClick(ByVal Item As MSComctlLib.ListItem)
'----------------------------------------------------------------------------------------'
Dim lIndex As Long

    lIndex = Val(Item.Tag)
    If lIndex = mlIndex Then Exit Sub
    
    mlIndex = lIndex

    Call QuestionSelect(Val(lIndex))

    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub picDrag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'----------------------------------------------------------------------------------------'
    mbDrag = True
End Sub

'----------------------------------------------------------------------------------------'
Private Sub picDrag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'----------------------------------------------------------------------------------------'
' calculate the proportioning of the two frames
'----------------------------------------------------------------------------------------'
'    If mbDrag Then
'        If Y < 100 Then Y = 100
'        If Y > Me.ScaleHeight - 100 Then Y = Me.ScaleHeight - 100
'
'        Y = Y + picDrag.Top
''        If (picDrag.Top > 30) And (Y < flxDiscrepancies.Top + flxDiscrepancies.Height - 480) Then
''        If (Y > 0) And (Y < Me.Height - picDrag.Height) Then
'            If (Y / (Me.ScaleHeight - cmdPrintListing.Height - 120) > 0) And (Y / (Me.ScaleHeight - cmdPrintListing.Height - 120) < 1) Then
'            'If (Y / (Me.ScaleHeight - cmdPrintListing.Height - 120) > 0.15) And (Y / (Me.ScaleHeight - cmdPrintListing.Height - 120) < 0.75) Then
'                picDrag.Top = Y
'                mdblProportion = picDrag.Top / (Me.ScaleHeight - cmdPrintListing.Height - 120)
'                Form_Resize
'                Me.Refresh
'            End If
' '       End If
'
'    End If
   If mbDrag Then
      Y = Y + picDrag.Top
      If Y > 0 And Y < Me.ScaleHeight - 1215 Then
            picDrag.Top = Y
           
            mdblProportion = picDrag.Top / (Me.ScaleHeight - cmdPrintListing.Height - 120)
           
            Form_Resize
            Me.Refresh
        End If
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub picDrag_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'----------------------------------------------------------------------------------------'
    mbDrag = False
    Form_Resize

End Sub

'----------------------------------------------------------------------------------------'
Private Function RecordSet_toListView(lvwlistview As MSComctlLib.ListView, rsRecordset As ADODB.Recordset, Optional sHeadings As String = "") As Long
'----------------------------------------------------------------------------------------'
' TA 15/05/2000
' uses row number as the listitem tag
' Ta 17/1/01: store widths in array and resize at end
'----------------------------------------------------------------------------------------'

Dim vHeadings As Variant
Dim lFields As Long
Dim i As Long
Dim sValue As String
Dim lRow As Long
Dim vWidth As Variant
    
    On Error GoTo ErrHandler
    
    lvwlistview.ListItems.Clear
    lvwlistview.ColumnHeaders.Clear
    
    lFields = rsRecordset.Fields.Count
    
    If sHeadings <> "" Then
        vHeadings = Split(sHeadings, "|")
    End If
    
    'set up array for widths
    ReDim vWidth(lFields - 1) As Long
    
    
    For i = 1 To lFields
        If IsArray(vHeadings) Then
            If i - 1 <= UBound(vHeadings) Then
                sValue = vHeadings(i - 1)
            Else
                sValue = rsRecordset.Fields(i - 1).Name
            End If
        Else
            sValue = rsRecordset.Fields(i - 1).Name
        End If
        lvwlistview.ColumnHeaders.Add , , sValue, lvwlistview.Parent.TextWidth(sValue) + 12 * Screen.TwipsPerPixelX
        'initialise width array
        vWidth(i - 1) = lvwlistview.Parent.TextWidth(sValue) + 12 * Screen.TwipsPerPixelX
    Next
    
    lRow = 0
    On Error GoTo ErrEmptyRS
    rsRecordset.MoveFirst
    On Error GoTo ErrHandler
    
    lRow = 1
     Do While Not rsRecordset.EOF
        sValue = RemoveNull(rsRecordset.Fields(0).Value)
        With lvwlistview.ListItems.Add(lRow, , sValue)
            .Tag = Format(lRow)
            If vWidth(0) < (lvwlistview.Parent.TextWidth(sValue) + 6 * Screen.TwipsPerPixelX) Then
                vWidth(0) = (lvwlistview.Parent.TextWidth(sValue) + 6 * Screen.TwipsPerPixelX)
            End If
            
            For i = 1 To lFields - 1
                sValue = RemoveNull(rsRecordset.Fields(i).Value)
                .SubItems(i) = sValue
                If vWidth(i) < (lvwlistview.Parent.TextWidth(sValue) + 12 * Screen.TwipsPerPixelX) Then
                  vWidth(i) = (lvwlistview.Parent.TextWidth(sValue) + 12 * Screen.TwipsPerPixelX)
            End If
            Next
        End With
        rsRecordset.MoveNext
        lRow = lRow + 1
     Loop
    
    'adjust column widths
    For i = 0 To lFields - 1
        lvwlistview.ColumnHeaders(i + 1).Width = vWidth(i)
    Next

    RecordSet_toListView = lRow - 1
    
Exit Function

ErrEmptyRS:

    RecordSet_toListView = 0

Exit Function
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "Recordset_toListview")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Function


'----------------------------------------------------------------------------------------'
Private Function RecordSet_toGrid(flxGrid As MSFlexGrid, rsRecordset As ADODB.Recordset, Optional sHeadings As String = "", Optional sColLengths As String = "", Optional bSizebyHeading As Boolean = True) As Long
'----------------------------------------------------------------------------------------'
' TA 19/05/2000
' uses row number as the listitem index
'----------------------------------------------------------------------------------------'
Dim vHeadings As Variant
Dim vColLengths As Variant
Dim vMinColLengths As Variant
Dim lFields As Long
Dim sValue As String
Dim lRow As Long
Dim nCol As Integer
Dim lColLength As Long
Dim lRowHeight As Long

    On Error GoTo ErrHandler
    
    lFields = rsRecordset.Fields.Count
    
    flxGrid.Clear
    flxGrid.Cols = lFields
    flxGrid.Rows = 1
   
    If sHeadings <> "" Then
        vHeadings = Split(sHeadings, "|")
    End If
        
    If sColLengths = "" Then
        sColLengths = String(lFields - 1, "|")
    End If
    vColLengths = Split(sColLengths, "|")
    
    vMinColLengths = vColLengths
    For nCol = 0 To lFields - 1
        If IsArray(vHeadings) Then
            If nCol <= UBound(vHeadings) Then
                sValue = vHeadings(nCol)
            Else
                sValue = rsRecordset.Fields(nCol).Name
            End If
        Else
            sValue = rsRecordset.Fields(nCol).Name
        End If
        flxGrid.TextMatrix(0, nCol) = sValue
        lColLength = Val(vColLengths(nCol))
        If lColLength = 0 Then
            If bSizebyHeading Then
                flxGrid.ColWidth(nCol) = (TextWidth(sValue) + 12 * Screen.TwipsPerPixelX)
            End If
        Else
            If bSizebyHeading Then
                If Len(sValue) < vMinColLengths(nCol) Then
                    'current value shorter than max
                    vMinColLengths(nCol) = Len(sValue)
                End If
            Else
                vMinColLengths(nCol) = 1
            End If
            flxGrid.ColWidth(nCol) = TextWidth(Left(sValue, vMinColLengths(nCol))) + (12 * Screen.TwipsPerPixelX)
        End If
    Next
    
    lRow = 0
    
    On Error GoTo ErrEmptyRS
    rsRecordset.MoveFirst
    On Error GoTo ErrHandler
    
    lRow = 1
     Do While Not rsRecordset.EOF
            sValue = ""
            For nCol = 0 To lFields - 1
                sValue = sValue & RemoveNull(rsRecordset.Fields(nCol).Value) & vbTab
            Next
        flxGrid.AddItem sValue
        rsRecordset.MoveNext
            
        With flxGrid
            .Row = .Rows - 1
            For nCol = 0 To lFields - 1
                .Col = nCol
                lColLength = Val(vColLengths(nCol))
                If lColLength = 0 Then
                    If .ColWidth(nCol) < (TextWidth(Trim(.Text)) + 12 * Screen.TwipsPerPixelX) Then
                        .ColWidth(nCol) = (TextWidth(Trim(.Text)) + 12 * Screen.TwipsPerPixelX)
                    End If
                Else
                    If vMinColLengths(nCol) <> lColLength Then
                        If Len(.Text) >= vColLengths(nCol) Then
                            vMinColLengths(nCol) = lColLength
                            flxGrid.ColWidth(nCol) = TextWidth(Left(.Text, lColLength)) + (12 * Screen.TwipsPerPixelX)
                        Else
                            If Len(.Text) > vMinColLengths(nCol) Then
                                If Len(.Text) < lColLength Then
                                    vMinColLengths(nCol) = Len(.Text)
                                Else
                                    vMinColLengths(nCol) = lColLength
                                End If
                                lColLength = vMinColLengths(nCol)
                                flxGrid.ColWidth(nCol) = TextWidth(Left(.Text, lColLength) & "00") + (12 * Screen.TwipsPerPixelX)
                            End If
                        End If
                    End If
                    lRowHeight = (TextWrapLines(.Text, lColLength) * TextHeight(.Text)) + (6 * Screen.TwipsPerPixelY)
                    If .RowHeight(.Row) < lRowHeight Then
                        .RowHeight(.Row) = lRowHeight
                    End If
                    .WordWrap = True
                End If
                .ColAlignment(nCol) = flexAlignLeftCenter
            Next
        End With
            
        
        lRow = lRow + 1
     Loop
     
    If flxGrid.Rows > 1 Then
       flxGrid.FixedRows = 1
    End If
     
    RecordSet_toGrid = lRow - 1
    
Exit Function

ErrEmptyRS:

    RecordSet_toGrid = 0
    
Exit Function

ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "Recordset_toGrid")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
        

End Function


'----------------------------------------------------------------------------------------'
Private Function TextWrapLines(ByVal sText As String, lCharLength As Long) As Long
'----------------------------------------------------------------------------------------'
' return numberof lines if text is wrapped at a certain character length
'----------------------------------------------------------------------------------------'
Dim lMarker As Long
Dim sChar As String
Dim sPortion As String
Dim lLines As Long

    sPortion = sText
    
    Do While sPortion <> ""
        
        For lMarker = lCharLength To 1 Step -1
            sChar = Mid(sPortion, lMarker, 1)
            If sChar = " " Then
                Exit For
            End If
        Next
        
        If lMarker = 0 Then
            lMarker = InStr(sPortion, " ")
            If lMarker = 0 Then
                lMarker = lCharLength
            End If
        End If
        sPortion = Mid(sPortion, lMarker + 1)
        lLines = lLines + 1
    Loop

    TextWrapLines = lLines
    
End Function


'---------------------------------------------------------------------
Private Sub View()
'---------------------------------------------------------------------
' View a particular SDV or Discrepancy
' NCJ 6 Nov 02 - Allow any status changes for SDVs
'---------------------------------------------------------------------
Dim oDisc As MIDiscrepancy
Dim oSDV As MISDV

    On Error GoTo ErrHandler
        
    cmdEdit.Enabled = False
    cmdReRaise.Enabled = False      ' "Queried" for SDVs
    cmdClose.Enabled = False        ' "Done" for SDVs
    cmdRespond.Enabled = False      ' "Cancel" for SDVs
    cmdPlanned.Enabled = False
    cmdShowDetails.Enabled = False
    lblPriority.TooltipText = ""
    
    If Not (moMIMsgs Is Nothing) Then
        cmdShowDetails.Enabled = True
        Select Case mnMIMsgType
        Case MIMsgType.mimtDiscrepancy
            Set oDisc = moMIMsgs
            If goUser.CheckPermission(gsFnCreateDiscrepancy) Then
                ' Responded discs can be re-raised
                If oDisc.CurrentMessage.Status = eDiscrepancyMIMStatus.dsResponded Then
                    cmdReRaise.Enabled = True
                End If
                ' Raised or Responded discs can be closed
                If oDisc.CurrentMessage.Status <> eDiscrepancyMIMStatus.dsClosed Then
                    cmdClose.Enabled = True
                End If
                ' Can only change priority of dicrepancies raised here
                If (oDisc.CurrentMessage.Status = eDiscrepancyMIMStatus.dsRaised) And (oDisc.CurrentMessage.Source = GetMIMsgSource) Then
                    lblPriority.TooltipText = "Right click to change priority"
                End If
                
            End If
            ' Need ChangeData permission to respond to dicrepancy
            If oDisc.CurrentMessage.Status = eDiscrepancyMIMStatus.dsRaised And goUser.CheckPermission(gsFnChangeData) Then
                cmdRespond.Enabled = True
            End If
        Case MIMsgType.mimtSDVMark
            Set oSDV = moMIMsgs
            ' Only users with CreateSDV permission can meddle with SDVs
            If goUser.CheckPermission(gsFnCreateSDV) Then
                cmdPlanned.Enabled = True
                cmdReRaise.Enabled = True
                cmdRespond.Enabled = True
                cmdClose.Enabled = True
                ' Disable the button corresponding to current status
                Select Case oSDV.CurrentMessage.Status
                Case eSDVMIMStatus.ssCancelled
                    cmdRespond.Enabled = False
                Case eSDVMIMStatus.ssDone
                    cmdClose.Enabled = False
                Case eSDVMIMStatus.ssPlanned
                    cmdPlanned.Enabled = False
                Case eSDVMIMStatus.ssQueried
                    cmdReRaise.Enabled = False
                End Select
            End If
        End Select
        
        ' edit message allowed if not sent yet and it's the same user (assume they have the same permissions!)
        If moMIMsgs.CurrentMessage.TimeSent = 0 And moMIMsgs.CurrentMessage.UserName = goUser.UserName Then
            cmdEdit.Enabled = True
        End If
       
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "View")
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
Private Sub PrintListingHeader(PrintingWidth As Long)
'---------------------------------------------------------------------
'Prints the Headings for the cmdPrintListing
'---------------------------------------------------------------------
Dim nCurrentY As Integer
Dim sHeaderLine As String

    On Error GoTo ErrHandler

    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.FontSize = 12
    Printer.FontBold = True
    Printer.Print "Macro - " & GetMIMTypeText(mnMIMsgType) & " Listing";
    
    Printer.FontBold = False
    Printer.FontSize = 8
    Printer.CurrentY = Printer.CurrentY + 60
    sHeaderLine = "Printed " & Format(Now, "yyyy/mm/dd hh:mm:ss") & "    Page " & Printer.Page
    Printer.CurrentX = PrintingWidth - Printer.TextWidth(sHeaderLine)
    Printer.Print sHeaderLine
    
    'draw a thicker line across page
    Printer.DrawWidth = 6
    Printer.CurrentY = Printer.CurrentY + 60
    nCurrentY = Printer.CurrentY
    Printer.Line (0, nCurrentY)-(PrintingWidth, nCurrentY)
    
    Printer.DrawWidth = 1
    Printer.CurrentY = Printer.CurrentY + 60
    nCurrentY = Printer.CurrentY
    'Heading banner from 0 inch to 10 1/4 inch
    Printer.Line (710, nCurrentY)-(PrintingWidth, nCurrentY + 240), , B
    Printer.CurrentY = nCurrentY + 30
    Printer.CurrentX = 720                                          '1/2 inch (720 twips)
    Printer.Print "Date/Time", ;                                    '1 3/8 inch (1980 twips)
    Printer.CurrentX = 2700                                         '(720+1980)
    Printer.Print "Status", ;                                       '5/8 inch (900 twips)
    If mnMIMsgType = MIMsgType.mimtDiscrepancy Then
        Printer.CurrentX = 3600                                     '(2700+900)
        Printer.Print "Priority", ;                                 '1/2 inch (720 twips)
    End If
    If mnMIMsgType = MIMsgType.mimtSDVMark Then
        Printer.CurrentX = 3600                                     '(2700+900)
        Printer.Print "Scope", ;                                 '1/2 inch (720 twips)
    End If
    Printer.CurrentX = 4320                                         '(3600+720)
    Printer.Print "Value", ;                                     '2 1/4 inch (3240 twips)
    Printer.CurrentX = 7560                                         '(4320+3240)
    Printer.Print "User", ;                                         '1 inch (1440 twips)
    Printer.CurrentX = 9000                                         '(7560+1440)
    Printer.Print "Message", ;
    Printer.CurrentY = nCurrentY + 270
    
Exit Sub
ErrHandler:
    'Changed 22/6/00 SR 3640
    Select Case Err.Number
        Case 482
            MsgBox "Printer error number 482 has occurred.", vbInformation, "MACRO"
        Case Else
            Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "PrintListingHeader")
                Case OnErrorAction.Ignore
                    Resume Next
                Case OnErrorAction.Retry
                    Resume
                Case OnErrorAction.QuitMACRO
                    Call ExitMACRO
                    Call MACROEnd
            End Select
    End Select
   
End Sub

'----------------------------------------------------------------------------------------'
Private Function CreatePrintingSQL(Optional bDCF As Boolean = False) As String
'----------------------------------------------------------------------------------------'
'When called from cmdPrintDCF_Click the optional paramater is set to true and the SQL statement
'is forced to only filter on status Raised.
'When called from cmdPrintListing_Click the optional parameter is not used and the SQL statement
'reflects the settings of chkRaised, chkResponded and chkClosed
'----------------------------------------------------------------------------------------'
'Mo 12/2/2003
'Changes to Order of this SQL stemming from SR 4800:-
'
'Original Order:-
'   MIMessageTrialName , MiMessageSite, TrialSubject.LocalIdentifier1, StudyVisit.VisitCode,
'   MIMessageVisitCycle, CRFPage.CRFPageCode, DataItemResponse.CRFPageCycleNumber, DataItem.DataItemCode
'New Order:-
'   MIMessageTrialName , MiMessageSite, TrialSubject.LocalIdentifier1, StudyVisit.VisitOrder,
'   MIMessageVisitCycle, CRFPage.CRFPageOrder, MIMessageCRFPageCycle, MIMessagePriority,
'   CRFElement.FieldOrder, MIMessageCreated
' DPH 19/01/2004 - Convert date to double when filtering on message created date in CreatePrintingSQL - SR5360
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim sIn As String
Const sSELECT_STATUS = "Please select at least one status"

    On Error GoTo ErrHandler
    
    'changed Mo Morris 24/11/00, DataItemResponse.LabResult and DataItemResponse.CTCGrade added
    sSQL = "SELECT MIMessageID, MIMessageSite, MIMessageSource, MIMessageType, MIMessageScope,"
    sSQL = sSQL & " MIMessageObjectID, MIMessageObjectSource, MIMessagePriority, MIMessageTrialName,"
    sSQL = sSQL & " MIMessagePersonId, MIMessageVisitId, MIMessageVisitCycle,"
    sSQL = sSQL & " MIMessageResponseTaskId, MIMessageResponseValue, MIMessageOCDiscrepancyID, MIMessageCreated,"
    sSQL = sSQL & " MIMessageSent, MIMessageReceived, MIMessageHistory, MIMessageProcessed,"
    sSQL = sSQL & " MIMessageStatus, MIMessageText, MIMessageUserName, MIMessageUserNameFull, MIMessageResponseTimeStamp,"
    sSQL = sSQL & " MIMessageResponseCycle, MIMessageCRFPageId, MIMessageCRFPageCycle, MIMessageDataItemId,"
    sSQL = sSQL & " TrialSubject.LocalIdentifier1,"
    sSQL = sSQL & " StudyVisit.VisitCode, StudyVisit.VisitName, StudyVisit.VisitOrder,"
    sSQL = sSQL & " DataItemResponse.LabResult, DataItemResponse.CTCGrade,"
    sSQL = sSQL & " CRFPage.CRFPageCode, CRFPage.CRFTitle, CRFPage.CRFPageOrder,"
    sSQL = sSQL & " DataItem.DataItemCode, DataItem.DataItemName, DataItem.DataType,"
    sSQL = sSQL & " CRFElement.FieldOrder"
    sSQL = sSQL & " FROM MIMessage, ClinicalTrial, TrialSubject, StudyVisit, DataItemResponse, CRFPage, DataItem, CRFElement"
    
    'join to table ClinicalTrial, to get the ClinicalTrialId which is needed to make other joins
    sSQL = sSQL & " WHERE MIMessage.MIMEssageTrialName = ClinicalTrial.ClinicalTrialName"
    
    'join to table TrialSubject, to get the LocalIdentifier1
    sSQL = sSQL & " AND ClinicalTrial.ClinicalTrialId = TrialSubject.ClinicalTrialId"
    sSQL = sSQL & " AND MIMessage.MIMessageSite = TrialSubject.TrialSite"
    sSQL = sSQL & " AND MIMessage.MIMessagePersonId = TrialSubject.PersonId"
    
    'join to table StudyVisit to get the VisitCode, VisitName and VisitOrder
    sSQL = sSQL & " AND ClinicalTrial.ClinicalTrialId = StudyVisit.ClinicalTrialId"
    sSQL = sSQL & " AND MIMessage.MIMessageVisitId = StudyVisit.VisitId"
    
    'join to table DataItemResponse to get LabResult and CTCGrade
    sSQL = sSQL & " AND ClinicalTrial.ClinicalTrialId = DataItemResponse.ClinicalTrialId"
    sSQL = sSQL & " AND MIMessage.MIMessageSite = DataItemResponse.TrialSite"
    sSQL = sSQL & " AND MIMessage.MIMessagePersonId = DataItemResponse.PersonId"
    sSQL = sSQL & " AND MIMessage.MIMessageResponseTaskID = DataItemResponse.ResponseTaskId"
    sSQL = sSQL & " AND MIMessage.MIMessageResponseCycle = DataItemResponse.RepeatNumber"
    
    'join to table CRFPage to get the CRFPageCode, CRFTitle and CRFPageOrder
    sSQL = sSQL & " AND ClinicalTrial.ClinicalTrialId = CRFPage.ClinicalTrialId"
    sSQL = sSQL & " AND MIMessage.MIMessageCRFPageId = CRFPage.CRFPageId"
    
    'join to table DataItem for the DataItemCode, DataItemName and DataType
    sSQL = sSQL & " AND ClinicalTrial.ClinicalTrialId = DataItem.ClinicalTrialId"
    sSQL = sSQL & " AND MIMessage.MIMessageDataItemId = DataItem.DataItemId"
    
    'join to table CRFElement for the FieldOrder
    sSQL = sSQL & " AND ClinicalTrial.ClinicalTrialId = CRFElement.ClinicalTrialId"
    sSQL = sSQL & " AND MIMessage.MIMessageCRFPageId = CRFElement.CRFPageId"
    sSQL = sSQL & " AND MIMessage.MIMessageDataItemId = CRFElement.DataItemId"
    
    'filtering on MessageType
    sSQL = sSQL & " AND MIMessageType = " & mnMIMsgType
    
    'filter for current messages only
    sSQL = sSQL & " AND MIMessageHistory = " & MIMsgHistory.mimhCurrent
    
    'Filter on site ?
    If mvParams(2) <> "" Then
        sSQL = sSQL & " AND MIMessageSite = '" & mvParams(2) & "'"
    End If

    'Filter on study/trial ?
    If mvParams(1) <> "" Then
        sSQL = sSQL & " AND MIMessageTrialName = '" & TrialNameFromId(CLng(mvParams(1))) & "'"
    End If

    'Filter on visit ?
    If mvParams(3) <> "" Then
        sSQL = sSQL & " AND MIMessageVisitId = " & mvParams(3)
    End If
    
    'Mo 9/6/2003, Bug 1836
    'Filter on eForm ?
    If mvParams(4) <> "" Then
        sSQL = sSQL & " AND MIMessageCRFPageId = " & mvParams(4)
    End If
    
    'Filter on Question ?
    If mvParams(5) <> "" Then
        sSQL = sSQL & " AND MIMessageDataItemId = " & mvParams(5)
    End If

    'Filter on subject label
    If mvParams(6) <> "" Then
        sSQL = sSQL & " AND " & GetSQLStringLike("LocalIdentifier1", mvParams(6))
    End If

'    'TA 30/10/2000: if we come though the view menu we need to restrict on user id
'    If msPerson <> "" Then
'        sSQL = sSQL & " AND MIMessagePersonId = " & msPerson
'    End If

    'Filter on creating user name
    If mvParams(7) <> "" Then
        sSQL = sSQL & " AND MIMessageUserName ='" & mvParams(7) & "'"
    End If
    
    'filter on message status mvParams(10)
    If bDCF Then
        sSQL = sSQL & " AND MIMEssageStatus =" & eDiscrepancyMIMStatus.dsRaised
    Else
        sIn = ""
        Select Case mnMIMsgType
        Case MIMsgType.mimtDiscrepancy
            'Discrepancy Status string contains Raised, Responded, Closed as 1 set or 0 not set
            If Mid(mvParams(10), 1, 1) = 1 Then
                sIn = sIn & eDiscrepancyMIMStatus.dsRaised & ","
            End If
            If Mid(mvParams(10), 2, 1) = 1 Then
                sIn = sIn & eDiscrepancyMIMStatus.dsResponded & ","
            End If
            If Mid(mvParams(10), 3, 1) = 1 Then
                sIn = sIn & eDiscrepancyMIMStatus.dsClosed & ","
            End If
            If sIn = "" Then
                DialogInformation sSELECT_STATUS, "Print Discrepancies"
                Exit Function
            End If
            'add statuses to where clause (knocking off final comma)
            sSQL = sSQL & " AND MIMEssageStatus IN (" & Left(sIn, Len(sIn) - 1) & ")"
        Case MIMsgType.mimtSDVMark
            'SDV Status string contains Planned, Done, Queried, Cancelled as 1 set or 0 not set
            If Mid(mvParams(10), 1, 1) = 1 Then
                sIn = sIn & eSDVMIMStatus.ssPlanned & ","
            End If
            If Mid(mvParams(10), 2, 1) = 1 Then
                sIn = sIn & eSDVMIMStatus.ssDone & ","
            End If
            If Mid(mvParams(10), 3, 1) = 1 Then
                sIn = sIn & eSDVMIMStatus.ssQueried & ","
            End If
            If Mid(mvParams(10), 4, 1) = 1 Then
                sIn = sIn & eSDVMIMStatus.ssCancelled & ","
            End If
            If sIn = "" Then
                DialogInformation sSELECT_STATUS, "Print SDV Marks"
                Exit Function
            End If
            'add statuses to where clause (knocking off final comma)
            sSQL = sSQL & " AND MIMEssageStatus IN (" & Left(sIn, Len(sIn) - 1) & ")"
        Case MIMsgType.mimtNote
            'Note Status string contains Public, Private as 1 set or 0 not set
            If Mid(mvParams(10), 1, 1) = 1 Then
                sIn = sIn & eNoteMIMStatus.nsPublic & ","
            End If
            If Mid(mvParams(10), 2, 1) = 1 Then
                sIn = sIn & eNoteMIMStatus.nsPrivate & ","
            End If
            If sIn = "" Then
                DialogInformation sSELECT_STATUS, "Print Notes"
                Exit Function
            End If
            'add statuses to where clause (knocking off final comma)
            sSQL = sSQL & " AND MIMEssageStatus IN (" & Left(sIn, Len(sIn) - 1) & ")"
        End Select
    End If
    
    'For SDVs filter on message scope mvParams(11)
    If mnMIMsgType = MIMsgType.mimtSDVMark Then
        sIn = ""
        'SDV Scope string contains Subject, Visit, eForm, Question as 1 set or 0 not set
        If Mid(mvParams(11), 1, 1) = 1 Then
            sIn = sIn & MIMsgScope.mimscSubject & ","
        End If
        If Mid(mvParams(11), 2, 1) = 1 Then
            sIn = sIn & MIMsgScope.mimscVisit & ","
        End If
        If Mid(mvParams(11), 3, 1) = 1 Then
            sIn = sIn & MIMsgScope.mimscEForm & ","
        End If
        If Mid(mvParams(11), 4, 1) = 1 Then
            sIn = sIn & MIMsgScope.mimscQuestion & ","
        End If
        'add SDV Scope to where clause (knocking off final comma)
        sSQL = sSQL & " AND MIMessageScope IN (" & Left(sIn, Len(sIn) - 1) & ")"
    End If

    'filter on message created date
    ' DPH 19/01/2004 - Convert date to double (CDbl(CDate(sDate))) for SQL - SR5360
    If mvParams(9) > "" Then
        ' DPH 19/01/2004 - check if mvParam(8) (before) = "true"
        If mvParams(8) = "true" Then
            sSQL = sSQL & " AND MIMessageCreated < " & LocalNumToStandard(CDbl(CDate(mvParams(9))))
        Else
            sSQL = sSQL & " AND MIMessageCreated > " & LocalNumToStandard(CDbl(CDate(mvParams(9))) + 1)
        End If
    End If

    'order the selection by TrialName/TrialSite/SubjectLabel/VisitOrder/VisitCycle/FormOrder/FormCycle/Priority/ElemeentOrder/Created
    sSQL = sSQL & " ORDER BY MIMessageTrialName, MiMessageSite, TrialSubject.LocalIdentifier1, "
    sSQL = sSQL & " StudyVisit.VisitOrder, MIMessageVisitCycle, CRFPage.CRFPageOrder, "
    sSQL = sSQL & " MIMessageCRFPageCycle, MIMessagePriority, CRFElement.FieldOrder, MIMessageCreated"

    CreatePrintingSQL = sSQL

Exit Function
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "CreatePrintingSQL")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Function

'---------------------------------------------------------------------
Private Function SplitMessageLine(ByRef sMessageLine As String, _
                                    nWidth As Integer, _
                                    nPrintFrom As Integer) As String
'---------------------------------------------------------------------
Dim i As Integer
Dim sPart As String
Dim sPreviousPart
Dim sChar As String
Dim q As Integer

    On Error GoTo ErrHandler

    'changed Mo Morris 12/2/01, SR 4071
    'to handle the manner in which this function works a space is added to the messageline
    'unless there is one already
    If Mid(sMessageLine, Len(sMessageLine), 1) <> " " Then
        sMessageLine = sMessageLine & " "
    End If
    sPart = ""
    sPreviousPart = ""
    For i = 1 To Len(sMessageLine)
        sChar = Mid(sMessageLine, i, 1)
        sPart = sPart & sChar
        If sChar = " " Then
            If Printer.TextWidth(sPart) > nWidth Then
                Printer.CurrentX = nPrintFrom
                'changed Mo Morris 12/2/01, SR 4071
                'Check for the situation where no spaces have been reached and the textwidth is beyond nWidth.
                'In this situation sPreviousPart would be empty and would need to have a truncated
                'section of sMessageLine placed in it
                If sPreviousPart = "" Then
                    q = Len(sMessageLine)
                    Do
                        q = q - 1
                        sPreviousPart = Mid(sMessageLine, 1, q)
                    Loop Until Printer.TextWidth(sPreviousPart) < nWidth
                End If
                Printer.Print sPreviousPart
                sMessageLine = Mid(sMessageLine, Len(sPreviousPart) + 1)
                Exit For
            Else
                sPreviousPart = sPart
            End If
        End If
    Next
          
    'check to see whether the remaining part of MessageLine requires a recursive call to SplitMessageLine
    If Printer.TextWidth(sMessageLine) < nWidth Then
        SplitMessageLine = sMessageLine
    Else
        SplitMessageLine = SplitMessageLine(sMessageLine, nWidth, nPrintFrom)
    End If
    
Exit Function
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "SplitMessageLine")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Function


'---------------------------------------------------------------------
Private Sub PrintDCFHeader(PrintingWidth As Long, _
                                sTrialName As String, _
                                sSite As String, _
                                sSubjectLabel As String)
'---------------------------------------------------------------------
'Prints the Headings for the cmdPrintDCF
'---------------------------------------------------------------------
Dim nCurrentY As Integer
Dim sHeaderLine As String

    On Error GoTo ErrHandler

    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.FontSize = 12
    Printer.FontBold = True
    Printer.Print "Macro - Data Clarification Form";
    
    Printer.FontBold = False
    Printer.FontSize = 8
    Printer.CurrentY = Printer.CurrentY + 60
    sHeaderLine = "Printed " & Format(Now, "yyyy/mm/dd hh:mm:ss") & "    Page " & Printer.Page
    Printer.CurrentX = PrintingWidth - Printer.TextWidth(sHeaderLine)
    Printer.Print sHeaderLine

    'draw a thicker line across page
    Printer.DrawWidth = 6
    Printer.CurrentY = Printer.CurrentY + 60
    nCurrentY = Printer.CurrentY
    Printer.Line (0, nCurrentY)-(PrintingWidth, nCurrentY)
    
    'print Trial/Site/Subject heading
    Printer.DrawWidth = 1
    Printer.CurrentY = Printer.CurrentY + 60
    Printer.FontBold = True
    Printer.CurrentX = 0
    Printer.Print "Study:", ;
    Printer.CurrentX = 1440
    'Changed Mo Morris 24/11/00
    Printer.Print TrialDescriptionFromName(sTrialName)
    'Printer.Print sTrialName & " - " & TrialDescriptionFromName(sTrialName)
    Printer.Print "Site:", ;
    Printer.CurrentX = 1440
    'Changed Mo Morris 24/11/00
    Printer.Print SiteDescriptionFromSite(sSite)
    'Printer.Print sSite & " - " & SiteDescriptionFromSite(sSite)
    Printer.Print "Subject:", ;
    Printer.CurrentX = 1440
    Printer.Print sSubjectLabel
    
    'draw a thicker line across page
    Printer.FontBold = False
    Printer.DrawWidth = 6
    Printer.CurrentY = Printer.CurrentY + 60
    nCurrentY = Printer.CurrentY
    Printer.Line (0, nCurrentY)-(PrintingWidth, nCurrentY)
    Printer.DrawWidth = 1
    
Exit Sub
ErrHandler:
    'Changed 22/6/00 SR 3640
    Select Case Err.Number
        Case 482
            MsgBox "Printer error number 482 has occurred.", vbInformation, "MACRO"
        Case Else
            Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "PrintDCFHeader")
                Case OnErrorAction.Ignore
                    Resume Next
                Case OnErrorAction.Retry
                    Resume
                Case OnErrorAction.QuitMACRO
                    Call ExitMACRO
                    Call MACROEnd
            End Select
    End Select
   
End Sub

'---------------------------------------------------------------------
Private Function TruncateStringForPrinting(ByRef sOriginalString As String, _
                                nWidth As Integer)
'---------------------------------------------------------------------
'This function will truncate a string at the last " " before Printer.TextWidth(String)
'exceeds the value in nWidth
'---------------------------------------------------------------------
Dim i As Integer
Dim sPart As String
Dim sPreviousPart As String
Dim sChar As String

    On Error GoTo ErrHandler
    
    'to handle the manner in which this function works a space is added to the messageline
    sOriginalString = sOriginalString & " "
    sPart = ""
    For i = 1 To Len(sOriginalString)
        sChar = Mid(sOriginalString, i, 1)
        sPart = sPart & sChar
        If sChar = " " Then
            If Printer.TextWidth(sPart) > nWidth Then
                TruncateStringForPrinting = sPreviousPart
                Exit For
            Else
                sPreviousPart = sPart
            End If
        End If
    Next
        
Exit Function
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "TruncateStringForPrinting")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Function

'----------------------------------------------------------------------------------------'
Private Function CycleNumberText(ByVal nCycleNumber As Integer) As String
'----------------------------------------------------------------------------------------'

    If nCycleNumber > 1 Then
        CycleNumberText = "[" & nCycleNumber & "]"
    End If

End Function

'----------------------------------------------------------------------------------------'
Private Sub GetResponseValueandTimeStamp(oMIMsgs As Object, ByRef sResponseValue As String, ByRef dblResponseTimeStamp As Double)
'----------------------------------------------------------------------------------------'
' Get the timestamp and response value for a specfic question.
' Note that sResponseValue and dblResponseTimeStamp as set by this sub.
'oMIMsgs is one of a MIDiscrepancy and MISDV
' NCJ 15 Oct 02 - Only get values if the Scope is Question
'----------------------------------------------------------------------------------------'
Dim oMIMsgList As MIDataLists
Dim vResponseInfo As Variant

    On Error GoTo Errlabel
    
    If oMIMsgs.Scope = MIMsgScope.mimscQuestion Then
        ' It's a question MIMsg
        Set oMIMsgList = New MIDataLists
        With oMIMsgs
            vResponseInfo = oMIMsgList.GetResponseDetails(gsADOConnectString, _
                                .StudyName, .Site, .SubjectId, .ResponseTaskId, .ResponseCycle)
        End With
        
        sResponseValue = ConvertFromNull(vResponseInfo(0, 0), vbString)
        dblResponseTimeStamp = vResponseInfo(1, 0)
        
        Set oMIMsgList = Nothing
    Else
        ' It doesn't have a Response
        sResponseValue = ""
        dblResponseTimeStamp = 0
    End If
    
Exit Sub
Errlabel:
    Err.Raise Err.Number, , Err.Description & "|frmViewDiscrepancies.GetResponseValueandTimeStamp"
    
End Sub
'----------------------------------------------------------------------------------------'
Private Function GetNRCTCTextFromResponseHistory(lClinicalTrialId As Long, _
                                                sTrialSite As String, _
                                                lPersonId As Long, _
                                                lResponseTaskId As Long, _
                                                dResponseTimeStamp As Double) As String
'----------------------------------------------------------------------------------------'
'   gets LabRersult and CTCGrade from DatItemResponseHistory
'   and then returns the result of calling GetNRCTCText
'   NCJ 20 Feb 01 - Deal with ResponseTimeStamp = 0
'   DPH 07/11/2003 - LocalNumToStandard date double used in SQL statement in GetNRCTCTextFromResponseHistory
'----------------------------------------------------------------------------------------'
Dim sSQL
Dim rsTemp As ADODB.Recordset
Dim sTemp As String

    On Error GoTo ErrHandler
    
    ' NCJ 20 Feb 01 - "Old" discrepancy data may not have a timestamp
    ' so ignore NTCTC for thes items
    If dResponseTimeStamp = 0 Then
        GetNRCTCTextFromResponseHistory = ""
        Exit Function
    End If
    
    ' DPH 07/11/2003 - LocalNumToStandard date double used in SQL statement in GetNRCTCTextFromResponseHistory
    sSQL = "SELECT DataItemResponseHistory.LabResult, DataItemResponseHistory.CTCGrade " _
        & " FROM DataItemResponseHistory " _
        & " WHERE DataItemResponseHistory.ClinicalTrialId = " & lClinicalTrialId _
        & " AND DataItemResponseHistory.TrialSite = '" & sTrialSite & "'" _
        & " AND DataItemResponseHistory.PersonId = " & lPersonId _
        & " AND DataItemResponseHistory.ResponseTaskId = " & lResponseTaskId _
        & " AND DataItemResponseHistory.ResponseTimeStamp = " & LocalNumToStandard(dResponseTimeStamp)

    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    sTemp = GetNRCTCText(rsTemp!labresult, rsTemp!CTCGrade)
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    GetNRCTCTextFromResponseHistory = sTemp
    
Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "GetNRCTCTextFromResponseHistory")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Function

'----------------------------------------------------------------------------------------'
Private Sub PrintDCFFooter()
'----------------------------------------------------------------------------------------'

    'On Error GoTo ErrHandler
    
    Printer.CurrentX = 0
    Printer.CurrentY = 10260    '7 1/8 inches
    Printer.FontSize = 10
    Printer.Print "Print Name: ______________________________   Signature: ______________________________   Date: ___/___/___"
    Printer.FontSize = 8


End Sub

'----------------------------------------------------------------------------------------'
Private Sub cmdPrintDCF_Click()
'----------------------------------------------------------------------------------------'
Dim rsMessages As ADODB.Recordset
Dim sSQL As String
Dim oMIMessage As clsMIMessage
Dim sPersonKey As String
Dim sPreviousPersonKey As String
Dim nYStorage As Long
Dim nExtraMessageLines As Integer
Dim sMessageText As String
Dim i As Long
Dim lPrintingWidth As Long
Dim nMessageWidth As Integer
Dim sResponseText As String
Dim sNRandCTCGrade As String
Dim sDataItemName As String

    On Error Resume Next

    'get the neccessary SQL statement
    HourglassOn
    sSQL = CreatePrintingSQL(True)

    Set rsMessages = New ADODB.Recordset
    rsMessages.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText

    If rsMessages.RecordCount = 0 Then
        HourglassOff
        Call DialogWarning("No raised discrepancies to print", "Print Data Clarification Form")
        Exit Sub
    End If

    'The Data Clarification Form only prints Discrepancies with a current status of raised
    'If the filters are not set to only display Raised discrepancies a message will be displayed
    If (Mid(mvParams(10), 1, 1) = 0) Or (Mid(mvParams(10), 2, 1) = 1) _
            Or (Mid(mvParams(10), 3, 1) = 1) Then
        Call DialogInformation("Only discrepancies with a current status of raised will be printed", _
                "Print Data Clarification Form")
    End If

    CommonDialog1.CancelError = True
    'Changed Mo 13/2/2003
    'Printer.Orientation = vbPRORLandscape
    CommonDialog1.Orientation = cdlLandscape

    Printer.TrackDefault = True
    CommonDialog1.ShowPrinter
    'Check for errors in ShowPrinter (including a cancel)
    If Err.Number > 0 Then
        HourglassOff
        Exit Sub
    End If

    'TA 30/10/2000: so we get the hourglass displayed
    DoEvents

    'restore normal error trapping
    On Error GoTo PrinterError

    'Changed Mo 13/2/2003
    Printer.Orientation = CommonDialog1.Orientation

    'set printer scalemode to twips
    Printer.ScaleMode = vbTwips
    Printer.FontSize = 8

    'extend the top and left margins by 1/2 inch to 3/4 inch approx
    Printer.ScaleLeft = -720
    Printer.ScaleTop = -720

    'Detext the printing width of the paper on the selected printer minus 2 * 1/2 inch borders (1440 twips)
    lPrintingWidth = Printer.ScaleWidth - 1440

    sPreviousPersonKey = ""
    'loop through each current message
    Do While Not rsMessages.EOF
        sPersonKey = rsMessages!MIMessageTrialName & "/" & rsMessages!MIMessageSite _
            & "/" & RemoveNull(rsMessages!LocalIdentifier1)
        If sPreviousPersonKey <> sPersonKey Then
            'throw a new page unless its the first subject
            If sPreviousPersonKey <> "" Then
                PrintDCFFooter
                Printer.NewPage
            End If
            sPreviousPersonKey = sPersonKey
            'each Subject is printed on a new page
            Call PrintDCFHeader(lPrintingWidth, rsMessages!MIMessageTrialName, rsMessages!MIMessageSite, _
                RemoveNull(rsMessages!LocalIdentifier1))
        End If

        'create a collection of the current message's history
        Set moMIMessages = New clsMIMessages
        moMIMessages.PopulateCollection rsMessages!MIMessageID, rsMessages!MIMessageSite, rsMessages!MIMessageObjectSource

        'check that there is enough space to print the current discrepancy and its history
        'based on 13 lines of 191 twips + 360 twips (2843) + number of history lines + extra lines for long text messages
        'an estimated extra line is added for every 65 characters of the message text
        'and a printing height of 10080 (7 inches)
        nExtraMessageLines = 0
        If RemoveNull(rsMessages!MIMessageText) > "" Then
            If Printer.TextWidth(rsMessages!MIMessageText) > 12960 Then
                nExtraMessageLines = nExtraMessageLines + 1
            End If
        End If
        For i = (moMIMessages.Count - 1) To 1 Step -1
            Set oMIMessage = moMIMessages.Item(i)
            nExtraMessageLines = nExtraMessageLines + 1 + Len(oMIMessage.MessageText) \ 65
        Next
        'Debug.Print "currentY=" & Printer.CurrentY & " space required=" & ((nExtraMessageLines * 191) + 2843) & " added=" & (Printer.CurrentY + ((nExtraMessageLines * 191) + 2843))
        If (Printer.CurrentY + (nExtraMessageLines * 191) + 2843) > 10080 Then
            PrintDCFFooter
            Printer.NewPage
            Call PrintDCFHeader(lPrintingWidth, rsMessages!MIMessageTrialName, rsMessages!MIMessageSite, _
                RemoveNull(rsMessages!LocalIdentifier1))
        End If

        'print the details of a discrepancy
        Printer.CurrentY = Printer.CurrentY + 60
        Printer.CurrentX = 0
        Printer.Print "Visit:", ;
        Printer.CurrentX = 1440
        Printer.Print rsMessages!VisitName & " [" & rsMessages!MIMessageVisitCycle & "]"
        Printer.Print "Form:", ;
        Printer.CurrentX = 1440
        Printer.Print rsMessages!CRFTitle & " [" & rsMessages!MIMessageCRFPageCycle & "]"
        Printer.Print "Question:", ;
        Printer.CurrentX = 1440
        
        sDataItemName = rsMessages!DataItemName
        If QuestionIsRQG(TrialIdFromName(rsMessages!MIMessageTrialName), rsMessages!MIMessageCRFPageId, rsMessages!MIMessageDataItemId) Then
             sDataItemName = sDataItemName & "[" & rsMessages!MIMessageResponseCycle & "]"
        End If
        
        Printer.Print sDataItemName
        Printer.Print "Raised by:", ;
        Printer.CurrentX = 1440
        Printer.Print rsMessages!MIMessageUserNameFull
        Printer.Print "Date/time:", ;
        Printer.CurrentX = 1440
        Printer.Print Format(rsMessages!MIMessageCreated, "yyyy/mm/dd hh:mm:ss"), ;
        Printer.CurrentX = 3600
        Printer.Print "Priority:", ;
        Printer.CurrentX = 4680
        Printer.Print rsMessages!MIMessagePriority
        Printer.Print "On Response:", ;
        Printer.CurrentX = 1440
        If rsMessages!DataType = DataType.LabTest Then
            sNRandCTCGrade = " [" & GetNRCTCText(rsMessages!labresult, rsMessages!CTCGrade) & "]"
        Else
            sNRandCTCGrade = ""
        End If
        Printer.Print rsMessages!MIMessageResponseValue & sNRandCTCGrade
        Printer.CurrentY = Printer.CurrentY + 120
        Printer.Print "Query:", ;
        sMessageText = RemoveNull(rsMessages!MIMessageText)
        nMessageWidth = lPrintingWidth - 1440                                           'Aprox 9 inch (12960 twips) on A4 paper
        If Printer.TextWidth(sMessageText) > nMessageWidth Then
            sMessageText = SplitMessageLine(sMessageText, nMessageWidth, 1440)
        End If
        Printer.CurrentX = 1440
        Printer.Print sMessageText
        Printer.CurrentY = Printer.CurrentY + 120
        Printer.Print "Answer:"
        Printer.Print ""
        nYStorage = Printer.CurrentY
        Printer.Line (1440, nYStorage)-(lPrintingWidth, nYStorage)
        Printer.Print ""
        Printer.Print ""
        nYStorage = Printer.CurrentY
        Printer.Line (1440, nYStorage)-(lPrintingWidth, nYStorage)
        Printer.Print ""
        Printer.Print ""
        nYStorage = Printer.CurrentY
        Printer.Line (1440, nYStorage)-(lPrintingWidth, nYStorage)

        'Print the discrepancy's history if there is one
        For i = (moMIMessages.Count - 1) To 1 Step -1
            If i = (moMIMessages.Count - 1) Then
                Printer.CurrentY = Printer.CurrentY + 60
                Printer.CurrentX = 0
                Printer.Print "History:", ;
            End If
            Set oMIMessage = moMIMessages.Item(i)
            Printer.CurrentX = 1440                                                     '1 inch (1440 twips)
            Printer.Print Format(oMIMessage.MessageCreated, "yyyy/mm/dd hh:mm:ss"), ;   '1 3/8 inch (1980 twips)
            Printer.CurrentX = 3420                                                     '(1440+1980)
            Printer.Print MACROMIMsgBS30.GetStatusText(mimtDiscrepancy, oMIMessage.MessageStatus), ;        '5/8 inch (900 twips)
            Printer.CurrentX = 4320                                                     '(3420+900)
            Printer.Print oMIMessage.MessagePriority, ;                                 '1/2 inch (720 twips)
            sResponseText = oMIMessage.MessageResponseValue
            If rsMessages!DataType = DataType.LabTest Then
                sNRandCTCGrade = GetNRCTCTextFromResponseHistory(TrialIdFromName(oMIMessage.MessageTrialName), _
                                        oMIMessage.MessageSite, oMIMessage.MessagePersonID, _
                                        oMIMessage.MessageResponseTaskId, oMIMessage.MessageResponseTimeStamp)
                sNRandCTCGrade = " [" & sNRandCTCGrade & "]"
            Else
                sNRandCTCGrade = ""
            End If
            sResponseText = sResponseText & sNRandCTCGrade
            If Printer.TextWidth(sResponseText) > 3240 Then
                sResponseText = TruncateStringForPrinting(sResponseText, 3240)
            End If
            Printer.CurrentX = 5040                                                     '(4320+720)
            Printer.Print sResponseText, ;                                              '2 1/4 inch (3240 twips)
            Printer.CurrentX = 8280                                                     '(5040+3240)
            Printer.Print oMIMessage.MessageUserName, ;                                 '1 inch (1440 twips)
            sMessageText = oMIMessage.MessageText
            nMessageWidth = lPrintingWidth - 9720                                       'Aprox 3 1/2 inch (5040 twips) on A4 paper
            If Printer.TextWidth(sMessageText) > nMessageWidth Then
                sMessageText = SplitMessageLine(sMessageText, nMessageWidth, 9720)
            End If
            Printer.CurrentX = 9720                                                    '(8280+1440)
            Printer.Print sMessageText
        Next

        'Print a line under the discrepancy
        Printer.CurrentY = Printer.CurrentY + 60
        nYStorage = Printer.CurrentY
        Printer.Line (0, nYStorage)-(lPrintingWidth, nYStorage)

        rsMessages.MoveNext
    Loop

    PrintDCFFooter
    Printer.EndDoc

    HourglassOff

Exit Sub
PrinterError:

    HourglassOff
    Call DialogError("A printer error has occurred.  The error number is " & Err.Number)

End Sub

'----------------------------------------------------------------------------------------'
Private Sub cmdPrintListing_Click()
'----------------------------------------------------------------------------------------'
'Note that this sub cannnot use the SQL statement created under cmdRefresh_Click, because
'it is not sorted in the order required by this print option. Creating the SQL statement
'here means that changes to the filters (made since the last refresh) will be reflected in
'the content of the listing.
'----------------------------------------------------------------------------------------'
Dim rsMessages As ADODB.Recordset
Dim sSQL As String
Dim oMIMessage As clsMIMessage
Dim sPersonKey As String
Dim sPreviousPersonKey As String
Dim sPersonHeading As String
Dim sQuestionKey As String
Dim sPreviousQuestionKey As String
Dim sQuestionHeading As String
Dim nHeadingWidth As Integer
Dim nYStorage As Long
Dim nExtraMessageLines As Integer
Dim sMessageText As String
Dim lPrintingWidth As Long
Dim nMessageWidth As Integer
Dim sResponseText As String
Dim sNRandCTCGrade As String
Dim sDataItemCode As String

    On Error Resume Next

    HourglassOn
    'get the neccessary SQL statement
    If mnMIMsgType <> MIMsgType.mimtSDVMark Then
        sSQL = CreatePrintingSQL
    Else
        sSQL = CreateSDVPrintingSQL
    End If

    Set rsMessages = New ADODB.Recordset
    rsMessages.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText

    If rsMessages.RecordCount = 0 Then
        HourglassOff
        Call DialogWarning("No discrepancies to print", "Print Listing")
        Exit Sub
    End If

    CommonDialog1.CancelError = True
    'Changed Mo 13/2/2003
    'Printer.Orientation = vbPRORLandscape
     CommonDialog1.Orientation = cdlLandscape
    Printer.TrackDefault = True
    CommonDialog1.ShowPrinter
    'Check for errors in ShowPrinter (including a cancel)
    If Err.Number > 0 Then
            HourglassOff
            Exit Sub
    End If

    'TA 30/10/2000: so we get the hourglass displayed
    DoEvents

    'restore normal error trapping
    On Error GoTo PrinterError
    
    'Changed Mo 13/2/2003
    Printer.Orientation = CommonDialog1.Orientation

    'set printer scalemode to twips
    Printer.ScaleMode = vbTwips
    Printer.FontSize = 8

    'extend the top and left margins by 1/2 inch to 3/4 inch approx
    Printer.ScaleLeft = -720
    Printer.ScaleTop = -720

    'Detext the printing width of the paper on the selected printer minus 2 * 1/2 inch borders (1440 twips)
    lPrintingWidth = Printer.ScaleWidth - 1440
    Call PrintListingHeader(lPrintingWidth)

    sPreviousPersonKey = ""
    sPreviousQuestionKey = ""
    'loop through each current message
    Do While Not rsMessages.EOF
        'create a collection of the current message's history
        Set moMIMessages = New clsMIMessages
        moMIMessages.PopulateCollection rsMessages!MIMessageID, rsMessages!MIMessageSite, rsMessages!MIMessageObjectSource
        'check that there is enough space to print the current message and its history
        'based on 190 twips per line, 460 twips for the 2 header lines and a printing height of 9780 (10080-300)
        'an estimated extra line is added for every 75 characters of the message text. (check the collection first)
        'i.e. 7 inches (7*1440 twips) - space required to print header line (300 twips)
        'search the message collections for message text that will require more than one line
        nExtraMessageLines = 0
        For Each oMIMessage In moMIMessages
            nExtraMessageLines = nExtraMessageLines + Len(oMIMessage.MessageText) \ 75
        Next
        If (Printer.CurrentY + ((moMIMessages.Count + nExtraMessageLines) * 190) + 460) > 9780 Then
            Printer.NewPage
            Call PrintListingHeader(lPrintingWidth)
            sPreviousPersonKey = ""
        End If

        sPersonKey = rsMessages!MIMessageTrialName & "/" & rsMessages!MIMessageSite _
            & "/" & RemoveNull(rsMessages!LocalIdentifier1)
        If sPreviousPersonKey <> sPersonKey Then
            'format and print Person Heading
            sPreviousPersonKey = sPersonKey
            sPreviousQuestionKey = ""
            sPersonHeading = "Study: " & rsMessages!MIMessageTrialName & "     Site: " _
                & rsMessages!MIMessageSite & "     Subject: " & RemoveNull(rsMessages!LocalIdentifier1)
            nYStorage = Printer.CurrentY
            Printer.FontBold = True
            nHeadingWidth = Printer.TextWidth(sPersonHeading)
            Printer.Line (-10, nYStorage)-(nHeadingWidth + 20, nYStorage + 240), , B
            Printer.CurrentY = nYStorage + 30
            Printer.CurrentX = 0
            Printer.Print sPersonHeading
            Printer.FontBold = False
            Printer.CurrentY = nYStorage + 270
        End If

        sDataItemCode = RemoveNull(rsMessages!DataItemCode)
        If QuestionIsRQG(TrialIdFromName(rsMessages!MIMessageTrialName), rsMessages!MIMessageCRFPageId, rsMessages!MIMessageDataItemId) Then
             sDataItemCode = sDataItemCode & "[" & rsMessages!MIMessageResponseCycle & "]"
        End If

        sQuestionKey = rsMessages!VisitCode & "/" & rsMessages!MIMessageVisitCycle & "/" & rsMessages!CRFPageCode _
            & "/" & rsMessages!MIMessageCRFPageCycle & "/" & sDataItemCode & "/" & rsMessages!MIMessageOCDiscrepancyID
        If sPreviousQuestionKey <> sQuestionKey Then
            'format and print Question Heading
            sPreviousQuestionKey = sQuestionKey
            'Note that for SDVs if the Scope is Subject there will be no VisitCode
            'Note that for SDVs if the Scope is Visit there will be no CRFPageCode
            'Note that for SDVs if the Scope is eForm there will be no DataItemCode
            If RemoveNull(rsMessages!VisitCode) = "" Then
                sQuestionHeading = "Visit: All Visits"
            ElseIf RemoveNull(rsMessages!CRFPageCode) = "" Then
                sQuestionHeading = "Visit: " & rsMessages!VisitCode & "[" & rsMessages!MIMessageVisitCycle & "]     Form: All eForms"
            ElseIf RemoveNull(rsMessages!DataItemCode) = "" Then
                sQuestionHeading = "Visit: " & rsMessages!VisitCode & "[" & rsMessages!MIMessageVisitCycle & "]     Form:" _
                    & rsMessages!CRFPageCode & "[" & rsMessages!MIMessageCRFPageCycle & "]     Question: All Questions"
            Else
                sQuestionHeading = "Visit: " & rsMessages!VisitCode & "[" & rsMessages!MIMessageVisitCycle & "]     Form:" _
                    & rsMessages!CRFPageCode & "[" & rsMessages!MIMessageCRFPageCycle & "]     Question:" _
                    & sDataItemCode
                If rsMessages!MIMessageOCDiscrepancyID <> 0 Then
                    sQuestionHeading = sQuestionHeading & "     Discrepancy Id:" & rsMessages!MIMessageOCDiscrepancyID
                End If
            End If
            nYStorage = Printer.CurrentY
            Printer.FontBold = True
            nHeadingWidth = Printer.TextWidth(sQuestionHeading)
            Printer.Line (350, nYStorage)-(nHeadingWidth + 380, nYStorage + 240), , B
            Printer.CurrentY = nYStorage + 30
            Printer.CurrentX = 360
            Printer.Print sQuestionHeading
            Printer.FontBold = False
            Printer.CurrentY = nYStorage + 270
        End If

        'print the repeating parts of a message
        For Each oMIMessage In moMIMessages
            Printer.CurrentX = 720                                                      '1/2 inch (720 twips)
            Printer.Print Format(oMIMessage.MessageCreated, "yyyy/mm/dd hh:mm:ss"), ;   '1 3/8 inch (1980 twips)
            Printer.CurrentX = 2700                                                     '(720+1980)
            Select Case mnMIMsgType
            Case MIMsgType.mimtDiscrepancy
                Printer.Print MACROMIMsgBS30.GetStatusText(mimtDiscrepancy, oMIMessage.MessageStatus), ;    '5/8 inch (900 twips)
            Case MIMsgType.mimtSDVMark
                Printer.Print MACROMIMsgBS30.GetStatusText(mimtSDVMark, oMIMessage.MessageStatus), ;
            Case MIMsgType.mimtNote
                Printer.Print MACROMIMsgBS30.GetStatusText(mimtNote, oMIMessage.MessageStatus), ;
            End Select
            If mnMIMsgType = MIMsgType.mimtDiscrepancy Then
                Printer.CurrentX = 3600                                                 '(2700+900)
                Printer.Print oMIMessage.MessagePriority, ;                             '1/2 inch (720 twips)
            End If
            If mnMIMsgType = MIMsgType.mimtSDVMark Then
                Printer.CurrentX = 3600                                                 '(2700+900)
                Select Case rsMessages!MIMessageScope
                Case MIMsgScope.mimscSubject
                    Printer.Print MACROMIMsgBS30.GetScopeText(MIMsgScope.mimscSubject), ;
                Case MIMsgScope.mimscVisit
                    Printer.Print MACROMIMsgBS30.GetScopeText(MIMsgScope.mimscVisit), ;
                Case MIMsgScope.mimscEForm
                    Printer.Print MACROMIMsgBS30.GetScopeText(MIMsgScope.mimscEForm), ;
                Case MIMsgScope.mimscQuestion
                    Printer.Print MACROMIMsgBS30.GetScopeText(MIMsgScope.mimscQuestion), ;
                End Select
            End If
            sResponseText = oMIMessage.MessageResponseValue
            'Changed Mo Morris 28/11/00 NormalRange+CTCGrdae added
            If rsMessages!DataType = DataType.LabTest Then
                sNRandCTCGrade = GetNRCTCTextFromResponseHistory(TrialIdFromName(oMIMessage.MessageTrialName), _
                                        oMIMessage.MessageSite, oMIMessage.MessagePersonID, _
                                        oMIMessage.MessageResponseTaskId, oMIMessage.MessageResponseTimeStamp)
                sNRandCTCGrade = " [" & sNRandCTCGrade & "]"
            Else
                sNRandCTCGrade = ""
            End If
            sResponseText = sResponseText & sNRandCTCGrade
            If Printer.TextWidth(sResponseText) > 3240 Then                             '2 1/4 inch (3240 twips)
                sResponseText = TruncateStringForPrinting(sResponseText, 3240)
            End If
            Printer.CurrentX = 4320                                                     '(3600+720)
            Printer.Print sResponseText, ;
            Printer.CurrentX = 7560                                                     '(4320+3240)
            Printer.Print oMIMessage.MessageUserName, ;                                 '1 inch (1440 twips)
            sMessageText = oMIMessage.MessageText
            nMessageWidth = lPrintingWidth - 9000                                       'Aprox 4 inch (5760 twips) on A4 paper
            If Printer.TextWidth(sMessageText) > nMessageWidth Then
                sMessageText = SplitMessageLine(sMessageText, nMessageWidth, 9000)
            End If
            Printer.CurrentX = 9000                                                     '(7560+1440)
            Printer.Print sMessageText
        Next
        rsMessages.MoveNext
    Loop

    Printer.EndDoc

    Call HourglassOff

Exit Sub
PrinterError:

    Call HourglassOff
    MsgBox "A printer error has occurred.  The error number is " & Err.Number, vbOKOnly + vbInformation

End Sub

'----------------------------------------------------------------------------------------'
Private Function IsValidString(sDescription As String) As Boolean
'----------------------------------------------------------------------------------------'
' Return TRUE if text is valid name for user
' Displays any necessary messages
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler
    
    IsValidString = False
    
    If sDescription > "" Then
        If Not gblnValidString(sDescription, valOnlySingleQuotes) Then
            MsgBox "A username" & gsCANNOT_CONTAIN_INVALID_CHARS, _
                    vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
        ElseIf Not gblnValidString(sDescription, valAlpha + valNumeric + valSpace) Then
            MsgBox " A username may only contain alphanumeric characters", _
                    vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
        ElseIf Len(sDescription) > 255 Then
            MsgBox " A username may not be more than 255 characters", _
                    vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
        Else
            IsValidString = True
        End If
    End If
    
    Exit Function
    
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "IsValidString")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Function

'----------------------------------------------------------------------------------------'
Private Sub RefreshOrCloseForm()
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    'TA 18/03/2003
    'update discrepancies and sdv count in task list
    frmMenu.UpdateDiscCount
    
    If mnLaunchMode = MonitorMIMessage Then
        If FormIsLoaded(g_DATAENTRY_FORM_NAME) Then
            'if eform loaded let's refresh the eForm ( update it's icon)
            frmEFormDataEntry.RefreshResponses
        End If
        'ASH 16/1/2003 Mimic refreshing of databrowser from righthand pane.
         Call Display(mvParams, mnMIMsgType, mnLaunchMode)
    Else
        Unload Me
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Private Function CreateSDVPrintingSQL() As String
'----------------------------------------------------------------------------------------'
'This sub is similar to CreatePrintingSQL (which is used for printing Discrepancies and Notes)
'CreateSDVPrintingSQL is specifically for printing SDVs.
'Because SDVs have a scope of Subject, Visit, eForm or Question the required data has to be
'collected by 4 separate SELECT statements joined by a UNION.
'The sort order in CreateSDVPrintingSQL is the same as in CreatePrintingSQL.
'When a sort field is not relevant (because of the SDVs scope) it is set to zero, as the
'following table shows this has the effect of sorting SDVs in the order of Subject/Visit/Form/Question:-
'
'SDV Scope  VisitOrder  VisitCycle  FormOrder   FormCycle   FieldOrder
'Subject        0           0           0           0           0
'Visit          set         set         0           0           0
'Form           set         set         set         set         0
'Question       set         set         set         set         set
'----------------------------------------------------------------------------------------'
Dim sSQLCommonSelect As String
Dim sSQLCommonWhere As String
Dim bPrevSELECT As Boolean
Dim sSQL As String
Dim sIn As String
Const sSELECT_STATUS = "Please select at least one status"

    On Error GoTo ErrHandler
    
    bPrevSELECT = False
    
    'Prepare the part of the SQL statement that will be used in all 4 Selects
    sSQLCommonSelect = "SELECT MIMessageID, MIMessageSite, MIMessageSource, MIMessageType, MIMessageScope," _
        & " MIMessageObjectID, MIMessageObjectSource, MIMessageTrialName," _
        & " MIMessagePersonId, MIMessageVisitId, MIMessageVisitCycle," _
        & " MIMessageResponseTaskId, MIMessageResponseValue, MIMessageOCDiscrepancyID, MIMessageCreated," _
        & " MIMessageSent, MIMessageReceived, MIMessageHistory, MIMessageProcessed," _
        & " MIMessageStatus, MIMessageText, MIMessageUserName, MIMessageUserNameFull, MIMessageResponseTimeStamp," _
        & " MIMessageResponseCycle, MIMessageCRFPageId, MIMessageCRFPageCycle, MIMessageDataItemId," _
        & " TrialSubject.LocalIdentifier1,"
        
    sSQLCommonWhere = " WHERE MIMessage.MIMEssageTrialName = ClinicalTrial.ClinicalTrialName" _
        & " AND ClinicalTrial.ClinicalTrialId = TrialSubject.ClinicalTrialId" _
        & " AND MIMessage.MIMessageSite = TrialSubject.TrialSite" _
        & " AND MIMessage.MIMessagePersonId = TrialSubject.PersonId"
    
    'Are SDVs with Scope Subject required
    'Mo 9/6/2003, Bug 1836, only collect Subject SDVs if no filtering on Visit, eForm and Question
    If (Mid(mvParams(11), 1, 1) = 1) _
        And (mvParams(3) = "") And (mvParams(4) = "") And (mvParams(5) = "") Then
        'Select SDVs with scope Subject
        sSQL = sSQLCommonSelect _
            & " '' as VisitCode, '' as VisitName, 0 as VisitOrder," _
            & " '' as LabResult, 0 as CTCGrade," _
            & " '' as CRFPageCode, '' as CRFTitle, 0 as CRFPageOrder," _
            & " '' as DataItemCode, '' as DataItemName, 0 as DataType," _
            & " 0 as FieldOrder" _
            & " FROM MIMessage, ClinicalTrial, TrialSubject"
        sSQL = sSQL & sSQLCommonWhere
        'Filter on Scope Subject
        sSQL = sSQL & " AND MIMessageScope = " & MIMsgScope.mimscSubject
        sSQL = sSQL & CreateSDVPrintingSQLFiltering
        bPrevSELECT = True
    End If
    
    'Are SDVs with Scope Visit required
    'Mo 9/6/2003, Bug 1836, only collect Visit SDVs if no filtering on eForm and Question
    If (Mid(mvParams(11), 2, 1) = 1) _
        And (mvParams(4) = "") And (mvParams(5) = "") Then
        If bPrevSELECT Then
            sSQL = sSQL & " UNION "
        End If
        'Select SDVs with scope Visit
        sSQL = sSQL & sSQLCommonSelect _
            & " StudyVisit.VisitCode, StudyVisit.VisitName, StudyVisit.VisitOrder," _
            & " '' as LabResult, 0 as CTCGrade," _
            & " '' as CRFPageCode, '' as CRFTitle, 0 as CRFPageOrder," _
            & " '' as DataItemCode, '' as DataItemName, 0 as DataType," _
            & " 0 as FieldOrder" _
            & " FROM MIMessage, ClinicalTrial, TrialSubject, StudyVisit"
        sSQL = sSQL & sSQLCommonWhere
        'join to table StudyVisit to get the VisitCode, VisitName and VisitOrder
        sSQL = sSQL & " AND ClinicalTrial.ClinicalTrialId = StudyVisit.ClinicalTrialId" _
            & " AND MIMessage.MIMessageVisitId = StudyVisit.VisitId"
        'Filter on Scope Visit
        sSQL = sSQL & " AND MIMessageScope = " & MIMsgScope.mimscVisit
        sSQL = sSQL & CreateSDVPrintingSQLFiltering
        'Filter on visit ?
        If mvParams(3) <> "" Then
            sSQL = sSQL & " AND MIMessageVisitId = " & mvParams(3)
        End If
        bPrevSELECT = True
    End If
    
    'Are SDVs with Scope eForm required
    'Mo 9/6/2003, Bug 1836, only collect eForm SDVs if no filtering on Question
    If (Mid(mvParams(11), 3, 1) = 1) _
        And (mvParams(5) = "") Then
        If bPrevSELECT Then
            sSQL = sSQL & " UNION "
        End If
        'Select SDVs with scope eForm
        sSQL = sSQL & sSQLCommonSelect _
            & " StudyVisit.VisitCode, StudyVisit.VisitName, StudyVisit.VisitOrder," _
            & " '' as LabResult, 0 as CTCGrade," _
            & " CRFPage.CRFPageCode, CRFPage.CRFTitle, CRFPage.CRFPageOrder," _
            & " '' as DataItemCode, '' as DataItemName, 0 as DataType," _
            & " 0 as FieldOrder" _
            & " FROM MIMessage, ClinicalTrial, TrialSubject, StudyVisit, CRFPage"
        sSQL = sSQL & sSQLCommonWhere
        'join to table StudyVisit to get the VisitCode, VisitName and VisitOrder
        sSQL = sSQL & " AND ClinicalTrial.ClinicalTrialId = StudyVisit.ClinicalTrialId" _
            & " AND MIMessage.MIMessageVisitId = StudyVisit.VisitId"
        'join to table CRFPage to get the CRFPageCode, CRFTitle and CRFPageOrder
        sSQL = sSQL & " AND ClinicalTrial.ClinicalTrialId = CRFPage.ClinicalTrialId" _
            & " AND MIMessage.MIMessageCRFPageId = CRFPage.CRFPageId"
        'Filter on Scope eForm
        sSQL = sSQL & " AND MIMessageScope = " & MIMsgScope.mimscEForm
        sSQL = sSQL & CreateSDVPrintingSQLFiltering
        'Filter on visit ?
        If mvParams(3) <> "" Then
            sSQL = sSQL & " AND MIMessageVisitId = " & mvParams(3)
        End If
        'Filter on eForm ?
        If mvParams(4) <> "" Then
            'Mo 9/6/2003, Bug 1836
            sSQL = sSQL & " AND MIMessageCRFPageId = " & mvParams(4)
        End If
        bPrevSELECT = True
    End If
    
    'Are SDVs with Scope Question required
    If Mid(mvParams(11), 4, 1) = 1 Then
        If bPrevSELECT Then
            sSQL = sSQL & " UNION "
        End If
        'Select SDVs with scope Question
        sSQL = sSQL & sSQLCommonSelect _
            & " StudyVisit.VisitCode, StudyVisit.VisitName, StudyVisit.VisitOrder," _
            & " DataItemResponse.LabResult, DataItemResponse.CTCGrade," _
            & " CRFPage.CRFPageCode, CRFPage.CRFTitle, CRFPage.CRFPageOrder," _
            & " DataItem.DataItemCode, DataItem.DataItemName, DataItem.DataType," _
            & " CRFElement.FieldOrder" _
            & " FROM MIMessage, ClinicalTrial, TrialSubject, StudyVisit, DataItemResponse, CRFPage, DataItem, CRFElement"
        sSQL = sSQL & sSQLCommonWhere
        'join to table StudyVisit to get the VisitCode, VisitName and VisitOrder
        sSQL = sSQL & " AND ClinicalTrial.ClinicalTrialId = StudyVisit.ClinicalTrialId" _
            & " AND MIMessage.MIMessageVisitId = StudyVisit.VisitId"
        'join to table CRFPage to get the CRFPageCode, CRFTitle and CRFPageOrder
        sSQL = sSQL & " AND ClinicalTrial.ClinicalTrialId = CRFPage.ClinicalTrialId" _
            & " AND MIMessage.MIMessageCRFPageId = CRFPage.CRFPageId"
        'join to table DataItemResponse to get LabResult and CTCGrade
        sSQL = sSQL & " AND ClinicalTrial.ClinicalTrialId = DataItemResponse.ClinicalTrialId" _
            & " AND MIMessage.MIMessageSite = DataItemResponse.TrialSite" _
            & " AND MIMessage.MIMessagePersonId = DataItemResponse.PersonId" _
            & " AND MIMessage.MIMessageResponseTaskID = DataItemResponse.ResponseTaskId" _
            & " AND MIMessage.MIMessageResponseCycle = DataItemResponse.RepeatNumber"
        'join to table DataItem for the DataItemCode, DataItemName and DataType
        sSQL = sSQL & " AND ClinicalTrial.ClinicalTrialId = DataItem.ClinicalTrialId" _
            & " AND DataItemResponse.DataItemId = DataItem.DataItemId"
        'join to table CRFElement for the FieldOrder
        sSQL = sSQL & " AND ClinicalTrial.ClinicalTrialId = CRFElement.ClinicalTrialId" _
            & " AND MIMessage.MIMessageCRFPageId = CRFElement.CRFPageId" _
            & " AND MIMessage.MIMessageDataItemId = CRFElement.DataItemId"
        'Filter on Scope Question
        sSQL = sSQL & " AND MIMessageScope = " & MIMsgScope.mimscQuestion
        sSQL = sSQL & CreateSDVPrintingSQLFiltering
        'Filter on visit ?
        If mvParams(3) <> "" Then
            sSQL = sSQL & " AND MIMessageVisitId = " & mvParams(3)
        End If
        'Filter on eForm ?
        If mvParams(4) <> "" Then
            'Mo 9/6/2003, Bug 1836
            sSQL = sSQL & " AND MIMessageCRFPageId = " & mvParams(4)
        End If
        'Filter on question ?
        If mvParams(5) <> "" Then
            'Mo 9/6/2003, Bug 1836
            sSQL = sSQL & " AND MIMessageDataItemId = " & mvParams(5)
        End If
        bPrevSELECT = True
    End If
    
    'order the selection by TrialName/TrialSite/SubjectLabel/VisitCode/VisitCycle/FormCode/FormCycle/DataItemCode
    sSQL = sSQL & " ORDER BY MIMessageTrialName, MiMessageSite, LocalIdentifier1, "
    sSQL = sSQL & " VisitOrder, MIMessageVisitCycle, CRFPageOrder, "
    sSQL = sSQL & " MIMessageCRFPageCycle, FieldOrder, MIMessageCreated"

    CreateSDVPrintingSQL = sSQL

Exit Function
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "CreateSDVPrintingSQL")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Function

'----------------------------------------------------------------------------------------'
Private Function QuestionIsRQG(lClinicalTrialId As Long, _
                                lCRFPageID As Long, _
                                lDataItemId As Long) As Boolean
'----------------------------------------------------------------------------------------'
'This function returns True if a Question is part of a RQG (Repeating Question Group)
'and False if it is not
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bTemp As Boolean

     On Error GoTo ErrHandler

    sSQL = "SELECT OwnerQGroupid FROM CRFElement " _
        & "WHERE ClinicalTrialId = " & lClinicalTrialId _
        & " AND CRFPageId = " & lCRFPageID _
        & " AND DataItemId = " & lDataItemId
        
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 1 Then
        If rsTemp!OwnerQGroupid > 0 Then
            bTemp = True
        Else
            bTemp = False
        End If
    Else
        bTemp = False
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    QuestionIsRQG = bTemp

Exit Function
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "QuestionIsRQG")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Function

'----------------------------------------------------------------------------------------'
Private Function CreateSDVPrintingSQLFiltering() As String
'----------------------------------------------------------------------------------------'
' REVISIONS
' DPH 19/01/2004 - Convert date to double when filtering on message created date in CreateSDVPrintingSQLFiltering - SR5360
'----------------------------------------------------------------------------------------'
Dim sIn As String
Dim sSQL As String
Const sSELECT_STATUS = "Please select at least one status"
    
    On Error GoTo ErrHandler
    
    'filtering on MessageType
    sSQL = sSQL & " AND MIMessageType = " & mnMIMsgType

    'filter for current messages only
    sSQL = sSQL & " AND MIMessageHistory = " & MIMsgHistory.mimhCurrent

    'Filter on site ?
    If mvParams(2) <> "" Then
        sSQL = sSQL & " AND MIMessageSite = '" & mvParams(2) & "'"
    End If

    'Filter on study/trial ?
    If mvParams(1) <> "" Then
        sSQL = sSQL & " AND MIMessageTrialName = '" & TrialNameFromId(CLng(mvParams(1))) & "'"
    End If

    'Filter on subject label
    If mvParams(6) <> "" Then
        sSQL = sSQL & " AND " & GetSQLStringLike("LocalIdentifier1", mvParams(6))
    End If

    'Filter on creating user name
    If mvParams(7) <> "" Then
        sSQL = sSQL & " AND MIMessageUserName ='" & mvParams(7) & "'"
    End If
    
    'filter on message status mvParams(10)
    sIn = ""
    'SDV Status string contains Planned, Done, Queried, Cancelled as 1 set or 0 not set
    If Mid(mvParams(10), 1, 1) = 1 Then
        sIn = sIn & eSDVMIMStatus.ssPlanned & ","
    End If
    If Mid(mvParams(10), 2, 1) = 1 Then
        sIn = sIn & eSDVMIMStatus.ssDone & ","
    End If
    If Mid(mvParams(10), 3, 1) = 1 Then
        sIn = sIn & eSDVMIMStatus.ssQueried & ","
    End If
    If Mid(mvParams(10), 4, 1) = 1 Then
        sIn = sIn & eSDVMIMStatus.ssCancelled & ","
    End If
    'add statuses to where clause (knocking off final comma)
    sSQL = sSQL & " AND MIMEssageStatus IN (" & Left(sIn, Len(sIn) - 1) & ")"
    
    'filter on message created date
    ' DPH 19/01/2004 - Convert date to double (CDbl(CDate(sDate))) for SQL - SR5360
    If mvParams(9) > "" Then
        ' DPH 19/01/2004 - check if mvParam(8) (before) = "true"
        If mvParams(8) = "true" Then
            sSQL = sSQL & " AND MIMessageCreated < " & LocalNumToStandard(CDbl(CDate(mvParams(9))))
        Else
            sSQL = sSQL & " AND MIMessageCreated > " & LocalNumToStandard(CDbl(CDate(mvParams(9))) + 1)
        End If
    End If
    
    CreateSDVPrintingSQLFiltering = sSQL

Exit Function
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "CreateSDVPrintingSQLFiltering")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Function


