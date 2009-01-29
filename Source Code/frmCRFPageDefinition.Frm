VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCRFPageDefinition 
   Caption         =   "eForm Definition"
   ClientHeight    =   9165
   ClientLeft      =   1200
   ClientTop       =   1860
   ClientWidth     =   10620
   Icon            =   "frmCRFPageDefinition.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9165
   ScaleWidth      =   10620
   Begin VB.TextBox txtDescription 
      Height          =   450
      Left            =   7590
      TabIndex        =   31
      Text            =   "txtDescription"
      Top             =   5340
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cboAuthorisationLevel 
      Height          =   315
      Left            =   6960
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   7380
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   10560
      TabIndex        =   11
      Top             =   8670
      Width           =   10620
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply"
         Height          =   295
         Left            =   2580
         TabIndex        =   17
         Top             =   60
         Width           =   1005
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   295
         Left            =   1200
         TabIndex        =   13
         Top             =   60
         Width           =   1005
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   295
         Left            =   120
         TabIndex        =   12
         Top             =   60
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdRenumber 
      Caption         =   "Renumber"
      Height          =   480
      Left            =   2580
      TabIndex        =   7
      Tag             =   "KeepBottomRight"
      Top             =   6720
      Width           =   1005
   End
   Begin VB.CommandButton cmdValidationExpression 
      BackColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   7560
      Picture         =   "frmCRFPageDefinition.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6600
      Width           =   600
   End
   Begin VB.CommandButton cmdMoveDown 
      Caption         =   "Move Down"
      Height          =   480
      Left            =   1200
      TabIndex        =   5
      Tag             =   "KeepBottom"
      Top             =   6720
      Width           =   1005
   End
   Begin VB.CommandButton cmdMoveUp 
      Caption         =   "Move Up"
      Height          =   480
      Left            =   120
      TabIndex        =   4
      Tag             =   "KeepBottom"
      Top             =   6720
      Width           =   1005
   End
   Begin VB.PictureBox picCRFPageDefinition 
      Align           =   1  'Align Top
      FillStyle       =   0  'Solid
      Height          =   7290
      Left            =   0
      ScaleHeight     =   7230
      ScaleWidth      =   10560
      TabIndex        =   0
      Top             =   0
      Width           =   10620
      Begin VB.OptionButton optLandscape 
         Caption         =   "Landscape"
         Height          =   195
         Left            =   8460
         TabIndex        =   30
         Top             =   660
         Width           =   1095
      End
      Begin VB.OptionButton optPortrait 
         Caption         =   "Portrait"
         Height          =   255
         Left            =   8460
         TabIndex        =   29
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox txtDisplayLength 
         Height          =   285
         Left            =   7560
         TabIndex        =   28
         Top             =   4680
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox cboFormDate 
         Height          =   315
         Left            =   4380
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   840
         Width           =   3195
      End
      Begin VB.CheckBox chkHideIfInactive 
         Caption         =   "Hide if inactive"
         Height          =   195
         Left            =   540
         TabIndex        =   26
         Top             =   780
         Width           =   1455
      End
      Begin VB.CheckBox chkDisplayNumbers 
         Caption         =   "Automatic numbering of questions"
         Height          =   195
         Left            =   540
         TabIndex        =   24
         Top             =   1035
         Width           =   3195
      End
      Begin VB.TextBox txtFormDate 
         Height          =   285
         Left            =   7440
         TabIndex        =   22
         Top             =   2640
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.CheckBox chkSequentialEntry 
         Caption         =   "Sequential Entry"
         Height          =   195
         Left            =   540
         TabIndex        =   20
         Top             =   1290
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox chkLocal 
         Alignment       =   1  'Right Justify
         Caption         =   "Local"
         Height          =   195
         Left            =   6760
         TabIndex        =   19
         Top             =   560
         Width           =   795
      End
      Begin MSFlexGridLib.MSFlexGrid flexCRFElement 
         Height          =   4995
         Left            =   120
         TabIndex        =   18
         Tag             =   "Resize"
         Top             =   1340
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   8811
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   3
      End
      Begin VB.CommandButton cmdFormLabelExpression 
         BackColor       =   &H00FFFFFF&
         Height          =   600
         Left            =   7680
         Picture         =   "frmCRFPageDefinition.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   60
         Width           =   600
      End
      Begin VB.ComboBox cboFormLabel 
         Height          =   315
         ItemData        =   "frmCRFPageDefinition.frx":0EDE
         Left            =   4380
         List            =   "frmCRFPageDefinition.frx":0EE0
         TabIndex        =   14
         Top             =   220
         Width           =   3195
      End
      Begin VB.TextBox txtCRFPageCode 
         Enabled         =   0   'False
         Height          =   315
         Left            =   540
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Text            =   "txtCRFPageCode"
         Top             =   60
         Width           =   1725
      End
      Begin VB.CheckBox chkMore 
         Caption         =   "Show conditions"
         Height          =   345
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   45
         Width           =   1320
      End
      Begin VB.TextBox txtCRFPageTitle 
         Height          =   315
         Left            =   540
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "txtCRFPageTitle"
         Top             =   420
         Width           =   3075
      End
      Begin VB.Label lblFormDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Date"
         Height          =   255
         Left            =   3800
         TabIndex        =   23
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblFormLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Label"
         Height          =   195
         Left            =   3900
         TabIndex        =   15
         Top             =   220
         Width           =   390
      End
      Begin VB.Label lblCRFPageCode 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblCRFPageTitle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Title"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   420
         Width           =   300
      End
   End
   Begin VB.TextBox txtExpression 
      Height          =   480
      Left            =   6960
      TabIndex        =   3
      Text            =   "txtExpression"
      Top             =   7740
      Width           =   3015
   End
   Begin VB.TextBox txtLabTestDateExpr 
      Height          =   480
      Left            =   6960
      TabIndex        =   27
      Text            =   "txtLabTestDateExpr"
      Top             =   8280
      Width           =   3015
   End
   Begin MSComctlLib.ImageList imglistLargeIcons 
      Left            =   4305
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imglistSmallIcons 
      Left            =   4830
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCRFPageDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998 - 2006. All Rights Reserved
'   File:       frmCRFPageDefinition.frm
'   Author:     Andrew Newbigging June 1997
'   Purpose:    Allows field order of CRF pages to be altered, or automatically
'   derived from the layout.  Also allows skip conditions to be set up.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
 '  1-18   Andrew Newbigging   4/06/97 - 27/8/98
'   19  Mo Morris           3/9/98  SR 420
'       Size of CheckBox button chkMore increased. chkMore_Click changed to resize
'       lsvCRFElement and picCRFPageDefinition in conjunction with the form size and
'       the state of chkMore. RefreshFields changed to handled the RecordCount=1
'       problem in the SQLServer version of Macro.
'       Mo Morris          13/10/98    SR 543
'       RefreshFields changed so that an exit is forced if there are no data items on
'       a form. The reason for this being that the Movefirst command which was added under
'       SPR 420 fails on an empy recordset.
'       Mo Morris           4/12/98 SR 432
'       lblCRFPageCode and txtCRFPageCode (readonly) added to this form for the purpose of
'       displaying a forms code alongside its editable title. Public Property Let CRFPageId
'       changed to enter CRFPageCode into txtCRFPageCode.
'       Mo Morris           10/12/98    SR 590
'       DataItemCode added to lsvCRFElement as subItem(2), with SkipCondition becoming subItem(3)
'       and CRFElementId becoming subItem(4). Like SkipCondition, DataItemCode is only visible
'       when the Show Conditions button has been pressed. Minor Changes made to RefreshFields,
'       cmdValidationExpression_Click, txtSkipCondition_Change, cmdApply_Click, cmdOK_Click
'   20  Paul Norris         06/08/99
'       Changed hard-coded flex grid column numbers to constants
'       Added new columns ReasonForChange, Mandatory
'       Added check box for control of sequential filling of a form
'       Added text box for default form date
'   NCJ 13/8/99 - Changed expression checking to Arezzo version
'   22  Paul Norris         06/08/99
'       Changed flex grid resizing back to original but with extended height( in chkMore_Click)
'   23  Paul Norris         06/09/99    SR1499  Added new field RoleCode to grid to
'                                               allow authorisation level to be set
'       Mo Morris   14/9/99     cmdFormDate_Click, cmdFormLabelExpression_Click and
'                               cmdValidationExpression_Click changed. Calls to
'                               frmExpressionbuilder replaced by call to frmCUIFunctionEditor
'   PN  28/09/99    Amended flexCRFElement_Click() to display empty string when required
'                   Added flexCRFElement_Scroll() to move editing controls during edits
'   PN  30/09/99    Amended Form_Load() to use clsTextList ADO security connection
'   NCJ 26 Oct 99   Removed DatabaseDescription from sSQL in Form Load
'   WillC 10/11/99  Added the error handlers
'   Mo Morris   16/11/99    DAO to ADO conversion
'   NCJ 30/11/99    Ensure Page Date Label is refreshed
'   MO Morris   10/12/99, cmdRenumber_Click changed to handle  not renumbering hidden dataitems
'   NCJ 11 Dec 99 Don't renumber pictures etc.
'   NCJ 13 Dec 99   Ids to Long
'   NCJ 15 Feb 00 Debugging changes to RoleCodes
'   NCJ 22 Feb 00 - SRs 2674 & 3086
'   TA 29/04/2000   subclassing removed
'   TA 14/09/2000: code for delaing with clinical test date expressions
'   TA 15/11/2000: resizing handling changes
'   TA 17/11/2000: renumbering code now in frmCRFDesign
'   NCJ 24/11/00 - Ensure question list is updated when changing eForm title
'   ZA 09/08/01 - Added date prompt for eForm
'   DPH 22/10/2001 - Added Check for click inside grid cell
'   NCJ 5 Dec 01 - Use frmCRFDesign.RefreshMe
'   NCJ 13 Dec 01 - Added extra columns to accommodate groups and ShowStatusFlag
'   NCJ 19 Dec 01 - Added group name and code to flexgrid columns
'   NCJ 7 Jan 02 - Tweaked column widths for better viewing
'   NCJ 15 Jan 02 - Use RefreshQuestionLists on frmMenu instead of private routine her
'   ZA 22/08/2002 - Added combo for setting form validation
'   ZA 18/09/2002 - Displaying Data Item Length and Display Length for a CRFElement
'   ASH 30/10/2002 - Changed "Data item length" to "Question Length" in Form_Load
'   ASH 5/11/2002 - Modified SaveFormDetails and Refreshfields to account for eform width
'   NCJ 22 Nov 02 - Disallow ShowFlag and RFC for Groups
'   ASH 13/01/2003 Re-Arranged order to move
'   'status flag','question length' and 'display length' to be next to 'only collect data if'
'   NCJ 28 Jan 03 - Removed unused chkPrompt and cmdFormDate
'   ASH 18 / 2 / 2003 - Added keypress_event for txtlabtestdateexpr and modified txtexp keypress_event.
'   NCJ 14 May 03 - Check lengths of AREZZO Collect if and Lab Test Date expressions
'   NCJ 15 May 03 - Reimplemented Move Up / Move Down buttons to take account of question groups (Bug 1344)
'   NCJ 15 May 03 - Reimplemented selection of eForm date which copes with moving questions (Bug 1721)
'   NCJ 30 Jun 03 - Bug 1900 - Reversed behaviour of "Move up"/"Move down" buttons
'   NCJ 1 Jul 03 - Only do Question/Display lengths for text box questions
'               Tidied up some validation, and don't validate when they click Cancel
'   NCJ 4 Sept 03 - Made Sequential Entry checkbox invisible (we don't support it any more)
'   ic 14/06/2005 added clinical coding
'   NCJ 15 May 06 - Consider eForm Access Mode
'   NCJ 21 Sept 06 - Call CUIFunctionEditor with appropriate Access Mode
'----------------------------------------------------------------------------------------'

Option Explicit
Option Compare Binary
Option Base 0

Private mbFormLoading As Boolean

Private mlCRFPageId As Long
Private msTmpFormat As String
Private mlClinicalTrialId As Long
Private mnVersionId As Integer
Private msClinicalTrialName As String
Private mnMouseDownOnRow As Integer

'constants for txtExpression indices
Private Const m_EXPRESSION_SKIPCONDITION = 0
Private Const m_EXPRESSION_LABTESTDATE = 1

' NCJ 13 Dec 01 - Added QGroup field order, Item Type and ShowFlag columns
'ASH 13/01/2003 Re-Arranged order to move status flag,
' question length and display length to be next to only collect data if
Private Const mCOL_FIELD_ORDER = 0
Private Const mCOL_QGFIELD_ORDER = 1
Private Const mCOL_DATA_ITEM_NAME = 2
Private Const mCOL_CODE = 3
Private Const mCOL_ITEM_TYPE = 4
Private Const mCOL_OPTIONAL = 5
Private Const mCOL_MANDATORY = 6
Private Const mCOL_HIDDEN = 7
Private Const mCOL_LOCAL = 8
Private Const mCOL_REASON_FOR_CHANGE = 9
Private Const mCOL_ONLY_COLLECT_IF = 10
Private Const mCOL_SHOWFLAG = 11
Private Const mCOL_DATA_ITEM_LENGTH = 12
Private Const mCOL_DISPLAY_LENGTH = 13
Private Const mCOL_CRFELEMENT_ID = 14
Private Const mCOL_AUTHORISATION_LEVEL = 15
Private Const mCOL_LABTESTDATE = 16
Private Const mCOL_QGROUPID = 17    ' NCJ 14 May 03
Private Const mCOL_ELEMENTUSE = 18    ' NCJ 15 May 03
Private Const mCOL_DESCRIPTION = 19    ' TA 18/06/2003


'ZA 03/09/2002 - increased from 15 to 17
' NCJ 14 May 03 - Increased from 17 to 18 (for QGROUPID)
' NCJ 15 May 03 - Increased from 18 to 19 (for ELEMENTUSE)
' TA 18 Jun 03 - Increased from 19 to 20 (for DESCRIPTION)
Private Const mCOL_COUNT = 20

Private Const msSELECTEDTEXT = "x"
Private Const msNOTSELECTEDTEXT = ""

Private mbAuthCodeWarningGiven As Boolean

'design time grid width
Private mlGridWidth As Long

' DPH 22/10/2001 Is click within the cell of grid?
Private mbClickInGridCell As Boolean

'ZA 21/08/2002 - keep track of the previously selected row for date type question
'dont' fire click event of Date question combo during the filling of the form
' NCJ 15 Mar 03 - Use msFormDateQuestion instead
'Private mnFormDateRow As Integer
Private msFormDateQuestion As String
Private mbIsGridLoading As Boolean

' NCJ 15 May 03
Private mnMaxFieldOrder As Integer
Private mnMinFieldOrder As Integer

' NCJ 15 May 06 - Access mode for the eForm
Private meAccessMode As eSDAccessMode

'------------------------------------------------------------------------------------'
Public Sub Display(lClinicalTrialId As Long, nVersionId As Integer, _
             sClinicalTrialName As String, lCRFPageID As Long, nAccessMode As eSDAccessMode)
'------------------------------------------------------------------------------------'
' Open the eForm definition dialog
' NCJ 15 May 06 - Added access mode parameter
'------------------------------------------------------------------------------------'

    On Error GoTo ErrHandler

    ' Set access mode first before we do anything!
    meAccessMode = nAccessMode
    Me.Icon = frmMenu.Icon
    
    Load Me
    
    ClinicalTrialId = lClinicalTrialId
    VersionId = nVersionId
    ClinicalTrialName = sClinicalTrialName
    ' Setting CRFPageId does lots of stuff...
    CRFPageId = lCRFPageID
    
    ' NCJ 20 Jun 06 - MUSD - Check we still exist!
    If Me.CRFPageId = 0 Then Exit Sub
    
    
    If meAccessMode = sdReadOnly Then cmdOK.Enabled = False
    
    Me.Show vbModal
 
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "Display", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If
End Sub

'---------------------------------------------------------------------
Public Property Get ClinicalTrialId() As Long
'---------------------------------------------------------------------

    ClinicalTrialId = mlClinicalTrialId

End Property

'---------------------------------------------------------------------
Public Property Let ClinicalTrialId(ByVal vClinicalTrialId As Long)
'---------------------------------------------------------------------

    mlClinicalTrialId = vClinicalTrialId

End Property

'---------------------------------------------------------------------
Public Property Get VersionId() As Integer
'---------------------------------------------------------------------

    VersionId = mnVersionId

End Property

'---------------------------------------------------------------------
Public Property Let VersionId(ByVal vVersionId As Integer)
'---------------------------------------------------------------------

    mnVersionId = vVersionId

End Property

'---------------------------------------------------------------------
Public Property Get ClinicalTrialName() As String
'---------------------------------------------------------------------

    ClinicalTrialName = msClinicalTrialName

End Property

'---------------------------------------------------------------------
Public Property Let ClinicalTrialName(ByVal vClinicalTrialName As String)
'---------------------------------------------------------------------

    msClinicalTrialName = vClinicalTrialName

End Property

'---------------------------------------------------------------------
Public Property Get CRFTitle() As String
'---------------------------------------------------------------------

    CRFTitle = txtCRFPageTitle.Text
    Me.Caption = txtCRFPageTitle.Text & " form definition"

End Property

'---------------------------------------------------------------------
Private Sub SetAuthorisationLevels()
'---------------------------------------------------------------------
Dim iStartRow As Integer
Dim iEndRow As Integer
Dim iRowIndex As Integer
Dim sStartEnd As String
    
    On Error GoTo ErrHandler
    
    sStartEnd = cboAuthorisationLevel.Tag
    If sStartEnd = "" Then  ' Tag not yet set up
        iStartRow = flexCRFElement.Row
        iEndRow = iStartRow
    Else
        iStartRow = ExtractFirstItemFromList(sStartEnd, ",")
        iEndRow = sStartEnd
    End If
    For iRowIndex = iStartRow To iEndRow
        flexCRFElement.TextMatrix(iRowIndex, mCOL_AUTHORISATION_LEVEL) = cboAuthorisationLevel
    Next
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "SetAuthorisationLevels")
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
Private Sub cboAuthorisationLevel_Click()
'---------------------------------------------------------------------
' PN change 23 - 03/09/99
On Error GoTo ErrHandler
    
    If Visible Then
        Call SetAuthorisationLevels
    End If
    
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cboAuthorisationLevel_Click")
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
Private Sub cboAuthorisationLevel_LostFocus()
'---------------------------------------------------------------------
' PN change 23 - 03/09/99
On Error GoTo ErrHandler
    
    Call SetAuthorisationLevels
    cboAuthorisationLevel.Enabled = False
    cboAuthorisationLevel.Visible = False
    
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cboAuthorisationLevel_LostFocus")
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
Private Sub cboFormDate_Click()
'---------------------------------------------------------------------
' They've selected this element as their Form Date question
'sets the appropriate values of Optional, Mandatory and Hidden values in the grid
'for the selected question
'---------------------------------------------------------------------
Dim nRow As Integer
    
'    nRow = cboFormDate.ItemData(cboFormDate.ListIndex)
    
    'if the question doesn't change , don't take any action
    'this event is also fired when this combo is being filled
'    If nRow = mnFormDateRow Or mbIsGridLoading Then Exit Sub
    If cboFormDate.Text = msFormDateQuestion Or mbIsGridLoading Then Exit Sub
    
    msFormDateQuestion = cboFormDate.Text
    
    With flexCRFElement
        'clear the previously selected row if there was any
        nRow = GetEFormDateRow
        If nRow > -1 Then
'        If mnFormDateRow > 0 And mnFormDateRow < .Rows Then
'            .Row = mnFormDateRow
            .Row = nRow
            .Col = mCOL_ELEMENTUSE
            .Text = eElementUse.User
            .Col = mCOL_OPTIONAL
            .Text = msNOTSELECTEDTEXT
        End If
            
        'set the values for the selected qeustion only if the selected question is not empty
'        If nRow > 0 Then
        If msFormDateQuestion > "" Then
        
            .Row = GetQuestionRow(msFormDateQuestion)
            .Col = mCOL_ELEMENTUSE
            .Text = eElementUse.EFormVisitDate
            .Col = mCOL_OPTIONAL
            .Text = msSELECTEDTEXT
            .Col = mCOL_MANDATORY
            .Text = msNOTSELECTEDTEXT
            .Col = mCOL_HIDDEN
            .Text = msNOTSELECTEDTEXT
            
        End If
        
    End With
    
'    mnFormDateRow = nRow
            
End Sub

'---------------------------------------------------------------------
Private Function GetEFormDateRow() As Integer
'---------------------------------------------------------------------
' Return the row number of the eFormDateElement (if any)
' Returns -1 if no eForm date element
'---------------------------------------------------------------------
Dim nRow As Integer
Dim nDateRow As Integer

    nDateRow = -1
    
    With flexCRFElement
        If .Rows > 1 Then
            For nRow = 1 To .Rows - 1
                If CLng(.TextMatrix(nRow, mCOL_ELEMENTUSE)) = eElementUse.EFormVisitDate Then
                    ' Found it
                    nDateRow = nRow
                    Exit For
                End If
            Next
        End If
    End With
    
    GetEFormDateRow = nDateRow
    
End Function

'---------------------------------------------------------------------
Private Function GetQuestionRow(sQuestion As String) As Integer
'---------------------------------------------------------------------
' Return the row number of the given question (assume it exists!)
'---------------------------------------------------------------------
Dim nRow As Integer
Dim nQRow As Integer

    With flexCRFElement
        For nRow = 1 To .Rows - 1
            If .TextMatrix(nRow, mCOL_CODE) = sQuestion Then
                ' Found it
                nQRow = nRow
                Exit For
            End If
        Next
    End With
    
    GetQuestionRow = nQRow
    
End Function

'---------------------------------------------------------------------
Private Sub cboFormLabel_Change()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If cboFormLabel.Text > "" Then
        chkLocal.Enabled = True
    Else
        chkLocal.Enabled = False
    End If
    
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cboFormLabel_Change")
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
Private Sub cboFormLabel_Click()
'---------------------------------------------------------------------
On Error GoTo ErrHandler
    
    If cboFormLabel.Text > "" Then
        chkLocal.Enabled = True
    Else
        chkLocal.Enabled = False
    End If
    
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cboFormLabel_Click")
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
Private Sub cboFormLabel_LostFocus()
'---------------------------------------------------------------------
' REVISIONS
' DPH 27/06/2005 - Bug 2480 check crf page label expression not greater than 255 characters
'---------------------------------------------------------------------
Dim sFormLabel As String

    On Error GoTo ErrHandler
    
    sFormLabel = Trim(cboFormLabel.Text)
     
    If Me.Visible = True Then
        If sFormLabel > "" Then
            If Not gblnValidString(sFormLabel, valOnlySingleQuotes) Then
                MsgBox "Form label" & gsCANNOT_CONTAIN_INVALID_CHARS, _
                        vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
                cboFormLabel.SetFocus
                Exit Sub
            ' Use Arezzo expression checker
            ElseIf Not gclmGuideline.IsValidExpression(sFormLabel) Then
                    MsgBox "Form label is not a valid expression.", _
                            vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
                    cboFormLabel.SetFocus
                    Exit Sub
            ' DPH 27/06/2005 - Bug 2480 check crf page label expression not greater than 255 characters
            ElseIf Len(sFormLabel) > 255 Then
                    MsgBox "Form label expression maximum length (255 characters) exceeded.", _
                            vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
                    cboFormLabel.SetFocus
                    Exit Sub
            End If
            chkLocal.Enabled = True
        Else
            chkLocal.Value = vbUnchecked
            chkLocal.Enabled = False
        End If
    End If
    
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cboFormLabel_LostFocus")
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
Private Sub cmdApply_Click()
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    SaveFormDetails
    
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "cmdApply_Click", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If
End Sub

'---------------------------------------------------------------------
Private Sub cmdCancel_Click()
'---------------------------------------------------------------------

    Unload Me

End Sub

'---------------------------------------------------------------------
Private Sub chkMore_Click()
'---------------------------------------------------------------------
'Changed by  Mo Morris   3/9/98      (change made in Released and Developed versions)
'Sizing changes to picCRFPageDefinition and lsvCRFElement.
'Note that chkMore_Click is now called from Form_Load
'---------------------------------------------------------------------
On Error GoTo ErrHandler
    
    

    ' PN change 22
    If chkMore.Value = False Then
        Me.Width = frmMenu.Width / 4
    Else
    'WillC SR2897 give greater width to allow viewing of requires authorisation column.
        Me.Width = 15000
        'Me.Width = 13000
    End If
    
    chkMore.Caption = "More..."
    cmdMoveUp.Top = Me.Height - Picture1.Height - 650 - (Me.Height - (Picture1.Top + Picture1.Height))
    cmdMoveDown.Top = cmdMoveUp.Top
    cmdRenumber.Top = cmdMoveUp.Top
    flexCRFElement.Width = picCRFPageDefinition.ScaleWidth - 60
    flexCRFElement.Height = cmdMoveUp.Top - flexCRFElement.Top - 200


Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "chkMore_Click")
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
Private Sub cmdFormLabelExpression_Click()
'---------------------------------------------------------------------
' Allow user to edit eForm label
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    frmCUIFunctionEditor.Initialise cboFormLabel, "Expression", _
                        (meAccessMode >= sdLayoutOnly), "Label for eForm " & Me.CRFTitle
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdFormLabelExpression_Click")
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
Private Function LastRowOfElement(ByVal nStartRow As Integer) As Integer
'---------------------------------------------------------------------
' NCJ 15 May 03
' What is the last grid row of the element which starts at nStartRow?
' For non-groups, it's the same row, otherwise include all the group's questions
' Assume that nStartRow does NOT refer to a group question
'---------------------------------------------------------------------
Dim nRow As Integer
Dim i As Integer

    On Error GoTo ErrLabel
    
    ' Default to current row
    nRow = nStartRow
    
    With flexCRFElement
        If CLng(.TextMatrix(nStartRow, mCOL_QGROUPID)) > 0 Then
            ' It's a group - see how many rows it takes up
            ' Start at the next row
            i = nStartRow + 1
            Do While i < .Rows
                If .TextMatrix(i, mCOL_QGFIELD_ORDER) > "" Then
                    ' It belongs to the group
                    nRow = nRow + 1
                    i = i + 1
                Else
                    ' We reached the last row of the group
                    Exit Do
                End If
            Loop
        End If
    End With
    
    LastRowOfElement = nRow

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmCRFPageDefinition.LastRowOfElement"

End Function

'---------------------------------------------------------------------
Private Function FirstRowOfElement(ByVal nEndRow As Integer) As Integer
'---------------------------------------------------------------------
' NCJ 15 May 03
' What is the first grid row of the element that ends at nEndRow?
' For non-groups, it's the same row, otherwise include all the group's questions
' nEndRow may refer to a question, a group or a group question
'---------------------------------------------------------------------
Dim nRow As Integer
Dim i As Integer

    On Error GoTo ErrLabel
    
    ' Default to current row
    nRow = nEndRow
    
    With flexCRFElement
        If CLng(.TextMatrix(nEndRow, mCOL_QGROUPID)) = 0 And .TextMatrix(nEndRow, mCOL_QGFIELD_ORDER) = "" Then
            ' It's not a group or a group question
        ElseIf CLng(.TextMatrix(nEndRow, mCOL_QGROUPID)) > 0 Then
            ' It's a group so assume it's only one row
            ' (because group questions are AHEAD of a group row)
        Else
            ' It's a group question - count back to its owning "group" row
            ' Start at the row below
            i = nEndRow - 1
            Do While i > 0
                nRow = nRow - 1
                If .TextMatrix(i, mCOL_QGFIELD_ORDER) > "" Then
                    ' We're still in the group
                    i = i - 1
                Else
                    ' It must be the group row itself
                    Exit Do
                End If
            Loop
        End If
    End With
    
    FirstRowOfElement = nRow

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmCRFPageDefinition.FirstRowOfElement"

End Function

'---------------------------------------------------------------------
Private Sub cmdMoveDown_Click()
'---------------------------------------------------------------------
' Move an element down one
' NCJ 15 May 03 - Rewritten to take account of question groups
' Assume that we CAN move this item down (button should not be enabled if we can't)
' NCJ 30 Jun 03 - "Down" means towards the bottom of the screen!!!
'---------------------------------------------------------------------
Dim nStartRow1 As Integer
Dim nStartRow2 As Integer
Dim nEndRow1 As Integer
Dim nEndRow2 As Integer

    On Error GoTo ErrHandler
    
    nStartRow1 = flexCRFElement.Row
    ' Find where this element finishes
    nEndRow1 = LastRowOfElement(nStartRow1)
    ' Now get the start and finish of the next element
    nStartRow2 = nEndRow1 + 1
    nEndRow2 = LastRowOfElement(nStartRow2)
    ' And swap them round!
    Call SwapElements(nStartRow1, nEndRow1, nEndRow2)

    ' Focus to where the element moved up to
    flexCRFElement.Row = nStartRow1 + (nEndRow2 - nStartRow2 + 1)
    flexCRFElement.Col = mCOL_FIELD_ORDER
    If Not flexCRFElement.RowIsVisible(flexCRFElement.Row) Then
        flexCRFElement.TopRow = flexCRFElement.Row
    End If

    ' Set the moveup/down buttons again
    Call EnableMoveUpMoveDown(flexCRFElement.Row)

    flexCRFElement.SetFocus
    

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "cmdMoveDown_Click", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdMoveUp_Click()
'---------------------------------------------------------------------
' Move an element up one
' NCJ 15 May 03 - Rewritten to take account of question groups
' Assume that we CAN move this item up (button should not be enabled if we can't)
' NCJ 30 Jun 03 - "Up" means towards the top of the screen!!!
'---------------------------------------------------------------------
Dim nStartRow1 As Integer
Dim nStartRow2 As Integer
Dim nEndRow1 As Integer
Dim nEndRow2 As Integer

    On Error GoTo ErrHandler
    
    ' This is the second element (to become the first)
    nStartRow2 = flexCRFElement.Row
    ' Find where this element finishes
    nEndRow2 = LastRowOfElement(nStartRow2)
    ' Find out what we're swapping with
    ' Previous element ends 1 row before this one starts
    nEndRow1 = nStartRow2 - 1
    ' Find where this element starts
    nStartRow1 = FirstRowOfElement(nEndRow1)
    ' And swap them round!
    Call SwapElements(nStartRow1, nEndRow1, nEndRow2)
    
    ' Focus to where the element moved down to
    flexCRFElement.Row = nStartRow1
    flexCRFElement.Col = mCOL_FIELD_ORDER
    If Not flexCRFElement.RowIsVisible(nStartRow1) Then
        flexCRFElement.TopRow = nStartRow1
    End If
    
    ' Set the moveup/down buttons again
    Call EnableMoveUpMoveDown(flexCRFElement.Row)
        
    flexCRFElement.SetFocus
    
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "cmdMoveUp_Click", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdOK_Click()
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    SaveFormDetails
    
    Unload Me
    
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "cmdOK_Click", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If
End Sub

'---------------------------------------------------------------------
Public Property Get CRFPageId() As Long
'---------------------------------------------------------------------

    CRFPageId = mlCRFPageId

End Property

'---------------------------------------------------------------------
Public Property Let CRFPageId(tmpCRFPageId As Long)
'---------------------------------------------------------------------
' Set the eForm ID for the eForm definition
' NCJ 20 Jun 06 - Check it still exists in MUSD
'---------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset

    Set rsTemp = gdsCRFPage(Me.ClinicalTrialId, Me.VersionId, tmpCRFPageId)
    If rsTemp.RecordCount <= 0 Then
        DialogInformation "This eForm no longer exists"
        rsTemp.Close
        Set rsTemp = Nothing
        mlCRFPageId = 0
        Exit Property
    End If
    
    ' NCJ 11/2/00 - Reset the "warning" flag when the CRFPageId changes
    If mlCRFPageId <> tmpCRFPageId Then
        mbAuthCodeWarningGiven = False
    End If
    
    mlCRFPageId = tmpCRFPageId
    txtCRFPageTitle.Text = rsTemp!CRFTitle
    txtCRFPageCode.Text = rsTemp!CRFPageCode
    RefreshFields (tmpCRFPageId)

End Property

'---------------------------------------------------------------------
Private Sub RefreshFields(lCRFPageID As Long)
'---------------------------------------------------------------------
'Changed by  Mo Morris   3/9/98      (change made in Released and Developed versions)
'The rsCRFElement.LoveLast and rsCRFElement.MoveFirst lines added because without them the SQLServer
'version of Macro was always returning a RecordCount of 1
'   Populates flexgrid
'   NCJ 30 Nov 99 - Also update form date label field
'   NCJ 11 Feb 00 SR2977 - Deal with non-existent role codes
'   ZA 21/08/2002 - add date questions to date combo for form date questions
'   ASH 4/11/2002 - added check for eform width
'   ic 14/06/2005 added clinical coding
'   NCJ 20 Jun 06 - MUSD - Check lCRFPageId > 0
'---------------------------------------------------------------------
Dim rsCRFElement As ADODB.Recordset
Dim rsCRFPage As ADODB.Recordset
Dim sRoleCode As String
Dim lDataItemId As Long
Dim lQGroupId As Long
Dim bHasGroups As Boolean
Dim oQGroup As QuestionGroup
Dim lWidth As Long

    On Error GoTo ErrHandler
    
    ' NCJ 20 Jun 06 - Check we have a valid CRFPageID
    If lCRFPageID = 0 Then Exit Sub
    
    'lsvCRFElement.ListItems.Clear
    cboFormLabel.Clear
    ' Initialise form date field - NCJ 30/11/99
    txtFormDate.Text = ""
    msFormDateQuestion = ""
    
    'clear previous question from the combo
    cboFormDate.Clear
'    mnFormDateRow = -1

    ' NCJ 22 Feb 00 - Reset Authorisation combo here
    Call RefreshAuthorisationCodes
    
    Set rsCRFPage = gdsCRFPage(Me.ClinicalTrialId, Me.VersionId, lCRFPageID)
    cboFormLabel.Text = RemoveNull(rsCRFPage!CRFPageLabel)
    chkLocal.Enabled = (cboFormLabel.Text > "")
    
    If rsCRFPage!LocalCRFPageLabel = 1 Then
        chkLocal.Value = vbChecked
    Else
        chkLocal.Value = vbUnchecked
    End If
    
    ' NCJ 15 May 06 - Don't use Sequential Entry any more
'    If rsCRFPage!SequentialEntry = 1 Then
'        chkSequentialEntry.Value = vbChecked
'    Else
'        chkSequentialEntry.Value = vbUnchecked
'    End If
    chkSequentialEntry.Value = vbUnchecked
    chkSequentialEntry.Enabled = False
    
    'SDM 06/01/00 SR2392
    If rsCRFPage!HideIfInactive = 1 Then
        chkHideIfInactive.Value = vbChecked
    Else
        chkHideIfInactive.Value = vbUnchecked
    End If
    
    ' NCJ 22 Feb 00 - SR2674
    If rsCRFPage!DisplayNumbers = 1 Then
        chkDisplayNumbers.Value = vbChecked
    Else
        chkDisplayNumbers.Value = vbUnchecked
    End If
    
    ' NCJ 30 Nov 99
    If Not IsNull(rsCRFPage!CRFPageDateLabel) Then
        txtFormDate.Text = rsCRFPage!CRFPageDateLabel
    End If
    
    ' ASH 4/11/2002
    If Not IsNull(rsCRFPage!EFormWidth) Then
        If rsCRFPage!EFormWidth = glLANDSCAPE_WIDTH Then
            optLandscape.Value = True
        Else
            optPortrait.Value = True
        End If
    Else
        optPortrait.Value = True
    End If
    
    rsCRFPage.Close
    Set rsCRFPage = Nothing
    
    Set rsCRFElement = gdsCRFPageDataItems(Me.ClinicalTrialId, Me.VersionId, lCRFPageID)
    
    ' See if the eForm has any groups
    bHasGroups = False
    mnMinFieldOrder = 999   ' Default to something big
    
    If rsCRFElement.RecordCount = 0 Then
    
        flexCRFElement.Rows = 1
        flexCRFElement.Row = 0
    
    Else
        
        'MoveLast and Movefirst make RecordCount return the correct value in the SQLServer version
        rsCRFElement.MoveLast
        msTmpFormat = Left$("0000000000", Len(Format((rsCRFElement.RecordCount))))
        
        flexCRFElement.Rows = 1
        flexCRFElement.Row = 0
        
        rsCRFElement.MoveFirst
        
        'ZA 21/08/2002 - add an empty row to form date combo
        cboFormDate.AddItem ""
        
        With flexCRFElement
            While Not rsCRFElement.EOF
                lDataItemId = rsCRFElement!DataItemId
                lQGroupId = rsCRFElement!QGroupID
                
                If lDataItemId > 0 Or lQGroupId > 0 Then
                    .Rows = .Rows + 1
                    If .Rows > 2 Then
                        .Row = .Row + 1
                    End If
                    .RowHeight(.Row) = cboAuthorisationLevel.Height
                    
                    ' NCJ 14 May 03 - Store QGroupId
                    .Col = mCOL_QGROUPID
                    .Text = lQGroupId
                
                    .Col = mCOL_CRFELEMENT_ID
                    .Text = rsCRFElement![CRFelementID]
                    
                    .Col = mCOL_FIELD_ORDER
                    .Text = Format(rsCRFElement![FieldOrder], msTmpFormat)
                    ' The max. field order will end up as the last one done
                    mnMaxFieldOrder = rsCRFElement![FieldOrder]
                    ' NCJ 30 Jun 03 - Keep track of min field order too
                    If rsCRFElement![FieldOrder] > 0 And rsCRFElement![FieldOrder] < mnMinFieldOrder Then
                        mnMinFieldOrder = rsCRFElement![FieldOrder]
                    End If
                    
                    ' NCJ 13 Dec 01 - Added QGroupFieldOrder
                    .Col = mCOL_QGFIELD_ORDER
                    If rsCRFElement![QGroupFieldOrder] > 0 Then
                        .Text = Format(rsCRFElement![QGroupFieldOrder], msTmpFormat)
                    Else
                        .Text = ""
                    End If
                    
                    ' NCJ 13/12/01 - Item type
                    .Col = mCOL_ITEM_TYPE
                    If lQGroupId > 0 Then
                        .Text = "Group"
                        bHasGroups = True
                        ' Assume frmMenu has the collection of Question Groups
                        Set oQGroup = frmMenu.QuestionGroups.GroupById(lQGroupId)
                    ElseIf rsCRFElement![OwnerQGroupID] > 0 Then
                        .Text = "Group Question"
                    Else
                        .Text = "Question"
                    End If
                    
                    If lDataItemId > 0 Then
                        .Col = mCOL_DATA_ITEM_NAME
                        .Text = rsCRFElement![DataItemName]
                        .Col = mCOL_CODE
                        .Text = rsCRFElement![DataItemCode]
                    Else
                        ' It's a group - get the group name & code
                        .Col = mCOL_DATA_ITEM_NAME
                        .Text = oQGroup.QGroupName
                        .Col = mCOL_CODE
                        .Text = oQGroup.QGroupCode
                    End If
                    
                    .Col = mCOL_OPTIONAL
                    .CellAlignment = flexAlignCenterCenter
                    If rsCRFElement!Optional = 0 Then
                        .Text = msNOTSELECTEDTEXT
                    Else
                        .Text = msSELECTEDTEXT
                    End If
                    .Col = mCOL_HIDDEN
                    .CellAlignment = flexAlignCenterCenter
                    If rsCRFElement!Hidden = 0 Then
                        .Text = msNOTSELECTEDTEXT
                    Else
                        .Text = msSELECTEDTEXT
                    End If
                    .Col = mCOL_LOCAL
                    .CellAlignment = flexAlignCenterCenter
                    ' ATO 31/7/01 Changed to LocalFlag
                    If rsCRFElement!LocalFlag = 0 Then
                        .Text = msNOTSELECTEDTEXT
                    Else
                        .Text = msSELECTEDTEXT
                    End If
                    .Col = mCOL_ONLY_COLLECT_IF
                    .CellAlignment = flexAlignLeftCenter
                    If Not IsNull(rsCRFElement![SkipCondition]) Then
                        .Text = rsCRFElement![SkipCondition]
                    End If
                    
                    ' PN change 20
                    ' Mandatory field
                    .Col = mCOL_MANDATORY
                    .CellAlignment = flexAlignCenterCenter
                    If rsCRFElement.Fields("Mandatory") = 1 Then
                        .Text = msSELECTEDTEXT
                    Else
                        .Text = msNOTSELECTEDTEXT
                    End If
                    ' reason for change field
                    .Col = mCOL_REASON_FOR_CHANGE
                    .CellAlignment = flexAlignCenterCenter
                    If rsCRFElement.Fields("RequireComment") = 1 Then
                        .Text = msSELECTEDTEXT
                    Else
                        .Text = msNOTSELECTEDTEXT
                    End If
                    
                    ' PN change - 03/09/99
                    ' new requires authorisation field
                    .Col = mCOL_AUTHORISATION_LEVEL
                    .Text = ""
                    ' get the name using the moAuthorisationLevels to map from the field value
                    ' NCJ SR2977 - Need to add this authorisation role if it doesn't exist
                    If Not IsNull(rsCRFElement.Fields("RoleCode")) Then
                        sRoleCode = rsCRFElement.Fields("RoleCode")
                        ' Trap error resulting from non-existent role code
                        ' NB "0" means no role code
                        If sRoleCode = "0" Then
                            sRoleCode = ""
                            .Text = ""
                        Else
                            On Error Resume Next
                            cboAuthorisationLevel.Text = sRoleCode
                            If Err.Number <> 0 Then
                                ' NCJ 11/2/00 - This RoleCode isn't in our Security DB
                                ' but we'll just add it to the list and the Combo
                                ' (NB It's not added to the Security DB)
                                'SDM 14/02/00 SR3000
                                'Call moAuthorisationLevels.AddItem(sRoleCode, sRoleCode)
                                cboAuthorisationLevel.AddItem (sRoleCode)
                                ' Set it to this one
                                cboAuthorisationLevel.ListIndex = cboAuthorisationLevel.NewIndex
                                ' Tell the user what we've done
                                Call WarnOfNonExistentRoleCode(sRoleCode)
                            End If
                            .Text = sRoleCode
                            ' Reset error handling
                            On Error GoTo ErrHandler
                        End If
                    End If
                    
                    ' Add data item codes to form label drop-down
                    If lDataItemId > 0 Then
                        cboFormLabel.AddItem rsCRFElement![DataItemCode]
                    End If
                    
                    'TA 14/09/2000: get lab test date expression
                    .Col = mCOL_LABTESTDATE
                    .CellAlignment = flexAlignLeftCenter
                    If Not IsNull(rsCRFElement![ClinicalTestDateExpr]) Then
                        .Text = rsCRFElement![ClinicalTestDateExpr]
                    End If
                                        
                    
                    ' NCJ 13/12/01 - Added ShowStatusFlag
                    .Col = mCOL_SHOWFLAG
                    .CellAlignment = flexAlignCenterCenter
                    ' Ensure we don't have a ShowFlag for groups
                    If rsCRFElement!ShowStatusFlag = 0 Or rsCRFElement!QGroupID > 0 Then
                        .Text = msNOTSELECTEDTEXT
                    Else
                        .Text = msSELECTEDTEXT
                    End If
                    
                    
                    'check if it is not a question group
                    If (rsCRFElement.Fields("QGroupID").Value = 0) Then
                    
                        .Col = mCOL_DATA_ITEM_LENGTH
                        .CellAlignment = flexAlignCenterCenter
                        'Length properties only apply to text box (not popup list - NCJ 1 Jul 03)
                        If (rsCRFElement.Fields("ControlType").Value = gn_TEXT_BOX) Then
                            .Text = RemoveNull(rsCRFElement.Fields("DataItemLength").Value)
                            
                            'ic 14/06/2005 clinical coding: thesaurus types cant have expand buttons
                            If (rsCRFElement.Fields("DataType").Value <> DataType.Thesaurus) Then
                                .Col = mCOL_DISPLAY_LENGTH
                                .CellAlignment = flexAlignCenterCenter
                                .Text = RemoveNull(rsCRFElement.Fields("DisplayLength").Value)
                            End If
                        End If
                        
                    End If
                
                    'ZA 21/08/2002 add date/time questions to date combo
                    'Don't add a question that is a part of question group
                    If rsCRFElement.Fields("DataType").Value = DataType.Date And rsCRFElement.Fields("OwnerQGroupID").Value = 0 Then
                        cboFormDate.AddItem rsCRFElement.Fields("DataItemCode").Value
                        ' cboFormDate.ItemData(cboFormDate.NewIndex) = .Row
                    End If
                    
                    ' NCJ 15 May 03 - Store ElementUse in special column
                    .Col = mCOL_ELEMENTUSE
                    .Text = rsCRFElement.Fields("ElementUse")
                    'ZA 21/08/2002 - check if the question is being used for form date
                    'if so, display it in the combo box
                    If rsCRFElement.Fields("ElementUse") = eElementUse.EFormVisitDate Then
                        cboFormDate.Text = rsCRFElement.Fields("DataItemCode")
                        msFormDateQuestion = rsCRFElement.Fields("DataItemCode")
                        ' mnFormDateRow = .Row
                    End If
                
                
                    'TA 18/06/2003: new description column
                    .Col = mCOL_DESCRIPTION
                    .CellAlignment = flexAlignLeftCenter
                    .Text = RemoveNull(rsCRFElement!Description)
                    'calculate width based on text passed in
                    lWidth = TextWidth(.Text & "    ")
                    If .ColWidth(mCOL_DESCRIPTION) < lWidth Then
                        .ColWidth(mCOL_DESCRIPTION) = lWidth
                    End If
                
                End If
                
                rsCRFElement.MoveNext   ' Move to next record.
            
            Wend
            
            If Not bHasGroups Then
                ' Hide the group columns
                .ColWidth(mCOL_QGFIELD_ORDER) = 0
                .ColWidth(mCOL_ITEM_TYPE) = 0
            End If
            
            .FixedCols = mCOL_ITEM_TYPE + 1
        
        End With
        
    End If
    
    rsCRFElement.Close
    Set rsCRFElement = Nothing
    
    'ZA 23/08/2002 - the grid has been filled so set the flag to true
    mbIsGridLoading = False
    
    cmdValidationExpression.Visible = False
    
    ' NCJ 13 Dec 01 - Temporarily disable the Up/Down buttons for groups
    ' (because they don't work yet!)
'    If bHasGroups Then
        cmdMoveUp.Enabled = False
        cmdMoveDown.Enabled = False
'    End If
    
    ' NCJ 15 May 06 - Set up things according to Access Mode
    Call SetFieldsAccess
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCRFPageDefinition.RefreshFields"
    
End Sub

'---------------------------------------------------------------------
Private Sub WarnOfNonExistentRoleCode(sRoleCode As String)
'---------------------------------------------------------------------
' Warn the user if a role code has been specified
' that doesn't exist in their Security DB
' Only give the warning once per form
'---------------------------------------------------------------------
Dim sMsg As String

    If Not mbAuthCodeWarningGiven Then
        sMsg = "The authorisation code" & vbCrLf
        sMsg = sMsg & sRoleCode & vbCrLf
        sMsg = sMsg & "specified for one or more questions on this form" & vbCrLf
        sMsg = sMsg & "does not exist in the current Security database."
        DialogWarning sMsg
        mbAuthCodeWarningGiven = True
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdRenumber_Click()
'---------------------------------------------------------------------
' Renumber all questions
'---------------------------------------------------------------------
Dim sMsg As String
    
    On Error GoTo ErrHandler
    
    sMsg = "Do you wish to save your changes to the eForm definition and renumber all questions to reflect their positions on the eForm?"
    
    If DialogQuestion(sMsg) = vbYes Then
        ' WillC 25/2/2000 SR2943
         'Save the details before doing the renumber otherwise you lose unsaved data
         Call SaveFormDetails
        
         Call frmCRFDesign.Renumber
         
         'ZA 23/08/2002 - set this flag to true until the grid is being refilled
         mbIsGridLoading = True
         
         'Refresh the Listview in frmCRFPageDEfinition based on new numbering
         RefreshFields (frmCRFDesign.CRFPageId)
         
    End If
    
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "cmdRenumber_Click", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdValidationExpression_Click()
'---------------------------------------------------------------------
'Bring up function editor form according to which column
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    If flexCRFElement.Col = mCOL_ONLY_COLLECT_IF Then
        ' "Collect if" condition
        If txtExpression.Visible Then
            frmCUIFunctionEditor.Initialise txtExpression, "Condition", _
                        (meAccessMode >= sdLayoutOnly), "Question's Collect if Condition"
        End If
    Else
        'lab test expression
        If txtLabTestDateExpr.Visible Then
            frmCUIFunctionEditor.Initialise txtLabTestDateExpr, "Expression", _
                        (meAccessMode >= sdLayoutOnly), "Laboratory Test Date Expression", "temporal"
         End If
    End If
    
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "cmdValidationExpression_Click", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub flexCRFElement_Scroll()
'---------------------------------------------------------------------
' PN 28/09/99 created to move controls when scrolling grid while
' editing is in progress
' if editing then save and stop editing
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    'SDM 21/12/99 SR2510
    'Hide the editing controls
    txtExpression.Visible = False
    txtExpression.Enabled = False
    txtLabTestDateExpr.Visible = False
    txtLabTestDateExpr.Enabled = False
    txtDisplayLength.Enabled = False
    txtDisplayLength.Visible = False
    txtDescription.Enabled = False
    txtDescription.Visible = False
    cmdValidationExpression.Visible = False
    DoEvents
    flexCRFElement.Refresh
    cboAuthorisationLevel.Visible = False
    
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "flexCRFElement_Scroll", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub flexCRFElement_Click()
'---------------------------------------------------------------------
' Click in grid to change an element property
' NCJ 22 Nov 02 - Disallow ShowFlag and RFC for Groups
' ic 14/06/2005 added clinical coding
' NCJ 15 May 06 - Disallow clicking if ReadOnly
'---------------------------------------------------------------------
Dim oElement As CRFElement

    On Error GoTo ErrHandler
    
    If meAccessMode = sdReadOnly Then Exit Sub
    
    Call EnableMoveUpMoveDown(flexCRFElement.Row)
    
    ' DPH 22/10/2001 - Do not do the following unless the Mouse click was within a cell of the grid
    If Not mbClickInGridCell Then
        Exit Sub
    End If
    
    With flexCRFElement
        If .Row > 0 Then
        
            ' Pick up the element from the current row
            Set oElement = frmCRFDesign.CRFElementById(.TextMatrix(.Row, mCOL_CRFELEMENT_ID))
            
            Select Case .Col
            ' PN change 20
            Case mCOL_MANDATORY, mCOL_OPTIONAL, mCOL_HIDDEN
                    
                'ZA 21/08/2002 - don't take any action if this is the eForm date question
'                If .Row <> mnFormDateRow Then
                If CLng(.TextMatrix(.Row, mCOL_ELEMENTUSE)) <> eElementUse.EFormVisitDate Then
                    ' toggle the x
                    If .Text = msSELECTEDTEXT Then
                        .Text = msNOTSELECTEDTEXT
                    Else
                        .Text = msSELECTEDTEXT
                        ' Keep Optional/Mandatory in step
                        If .Col = mCOL_MANDATORY Then
                            .TextMatrix(.Row, mCOL_OPTIONAL) = msNOTSELECTEDTEXT
                        ElseIf .Col = mCOL_OPTIONAL Then
                            .TextMatrix(.Row, mCOL_MANDATORY) = msNOTSELECTEDTEXT
                        End If
                    End If
                End If
            
            Case mCOL_LOCAL
                    ' toggle the x
                    If .Text = msSELECTEDTEXT Then
                            .Text = msNOTSELECTEDTEXT
                    Else
                        .Text = msSELECTEDTEXT
                        ' Keep Optional/Mandatory in step
                        If .Col = mCOL_MANDATORY Then
                            .TextMatrix(.Row, mCOL_OPTIONAL) = msNOTSELECTEDTEXT
                        ElseIf .Col = mCOL_OPTIONAL Then
                            .TextMatrix(.Row, mCOL_MANDATORY) = msNOTSELECTEDTEXT
                        End If
                    End If
                
            Case mCOL_REASON_FOR_CHANGE, mCOL_SHOWFLAG
                    ' toggle the x if not a Group
                    If oElement.QGroupID = 0 Then
                        If .Text = msSELECTEDTEXT Then
                                .Text = msNOTSELECTEDTEXT
                        Else
                            .Text = msSELECTEDTEXT
                            ' Keep Optional/Mandatory in step
                            If .Col = mCOL_MANDATORY Then
                                .TextMatrix(.Row, mCOL_OPTIONAL) = msNOTSELECTEDTEXT
                            ElseIf .Col = mCOL_OPTIONAL Then
                                .TextMatrix(.Row, mCOL_MANDATORY) = msNOTSELECTEDTEXT
                            End If
                        End If
                    End If
            
            Case mCOL_ONLY_COLLECT_IF
                'ZA 21/08/2002 - don't take any action if this question is selected for form validation
'                If .Row <> mnFormDateRow Then
                If CLng(.TextMatrix(.Row, mCOL_ELEMENTUSE)) <> eElementUse.EFormVisitDate Then
                
                    ' edit the field and display the function builder button
                    txtExpression.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
                    txtExpression.Text = .Text
                    If .Row <= .RowSel Then
                        txtExpression.Tag = .Row & "," & .RowSel
                    Else
                        txtExpression.Tag = .RowSel & "," & .Row
                    End If
                    txtExpression.Enabled = True
                    txtExpression.SelStart = 0
                    txtExpression.SelLength = Len(txtExpression.Text)
                    txtExpression.Visible = True
                    txtExpression.ZOrder
                    
                    cmdValidationExpression.Move .Left + .CellLeft + .CellWidth - cmdValidationExpression.Width, _
                        .Top + .CellTop - (cmdValidationExpression.Height - .CellHeight) / 2
                    cmdValidationExpression.Visible = True
                    cmdValidationExpression.ZOrder
                    txtExpression.SetFocus
                End If
    
            Case mCOL_LABTESTDATE
                If oElement.DataItemType = DataType.LabTest Then
                    'only does this if it's the skip condition or we're dealing with a lab test question
                    ' edit the field and display the function builder button
                    txtLabTestDateExpr.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
                    txtLabTestDateExpr.Text = .Text
                    'currently only the first selected cell will be changed
                    txtLabTestDateExpr.Tag = .Row & "," & .Row
                    txtLabTestDateExpr.Enabled = True
                    txtLabTestDateExpr.SelStart = 0
                    txtLabTestDateExpr.SelLength = Len(txtLabTestDateExpr.Text)
                    txtLabTestDateExpr.Visible = True
                    txtLabTestDateExpr.ZOrder
                    
                    cmdValidationExpression.Move .Left + .CellLeft + .CellWidth - cmdValidationExpression.Width, _
                        .Top + .CellTop - (cmdValidationExpression.Height - .CellHeight) / 2
                    cmdValidationExpression.Visible = True
                    cmdValidationExpression.ZOrder
                    txtLabTestDateExpr.SetFocus
                End If
                
            Case mCOL_AUTHORISATION_LEVEL
                'authorisation Column Is Clicked - only for non-groups
                If oElement.QGroupID = 0 Then
                    If .Row <= .RowSel Then
                        cboAuthorisationLevel.Tag = .Row & "," & .RowSel
                    Else
                        cboAuthorisationLevel.Tag = .RowSel & "," & .Row
                    End If
                    
                    With cboAuthorisationLevel
                        .Top = flexCRFElement.CellTop + flexCRFElement.Top + picCRFPageDefinition.Top
                        .Left = flexCRFElement.CellLeft + flexCRFElement.Left + picCRFPageDefinition.Left
                        .Width = flexCRFElement.CellWidth
                        .Enabled = True
                        If flexCRFElement.Text <> vbNullString Then
                            .Text = flexCRFElement.Text
                        
                        ' PN 28/09/99 if the text is empty string then set the combo item to 0
                        Else
                            .ListIndex = 0
                        
                        End If
                        
                        .Visible = True
                        ' bring to front
                        .ZOrder
                    End With
                End If
                
            Case mCOL_DISPLAY_LENGTH
                'ZA 13/09/2002 - allow user to edit display length property only for text/popup list
                ' NCJ 1 Jul 03 - Only for text boxes
                Select Case oElement.ControlType
                Case gn_TEXT_BOX
                    'ic 14/06/2005 clinical coding: thesaurus types cant have expand buttons
                    If (oElement.DataItemType <> DataType.Thesaurus) Then

                        txtDisplayLength.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
                        txtDisplayLength.Text = .Text
                        
                        txtDisplayLength.Tag = .Row & "," & .Row
                        txtDisplayLength.Enabled = True
                        txtDisplayLength.SelStart = 0
                        txtDisplayLength.SelLength = Len(txtDisplayLength.Text)
                        txtDisplayLength.Visible = True
                        txtDisplayLength.ZOrder
                        txtDisplayLength.SetFocus
                    End If
                End Select
                
            Case mCOL_DESCRIPTION
                'TA 16/06/2003 - allow user to edit description

                txtDescription.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
                txtDescription.Text = .Text
                
                txtDescription.Tag = .Row & "," & .Row
                txtDescription.Enabled = True
                txtDescription.SelStart = 0
                txtDescription.SelLength = Len(txtDescription.Text)
                txtDescription.Visible = True
                txtDescription.ZOrder
                txtDescription.SetFocus

                
            End Select
        End If
    End With
    
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "flexCRFElement_Click", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub flexCRFElement_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
' MLM 26/03/03: Bug 1343: Set mbClickInGridCell = False if they clicked in a header cell
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    mnMouseDownOnRow = flexCRFElement.Row
    
    ' DPH 22/10/2001 - Is click within cell?
    If X > flexCRFElement.ColPos(flexCRFElement.Col) + flexCRFElement.ColWidth(flexCRFElement.Col) _
        Or X < flexCRFElement.ColPos(mCOL_ITEM_TYPE) _
        Or Y > flexCRFElement.RowPos(flexCRFElement.Row) + flexCRFElement.RowHeight(flexCRFElement.Row) _
        Or Y < flexCRFElement.RowPos(0) + flexCRFElement.RowHeight(0) Then
        mbClickInGridCell = False
    Else
        mbClickInGridCell = True
    End If

Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "flexCRFElement_MouseDown", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub RefreshAuthorisationCodes()
'---------------------------------------------------------------------
' NCJ 22 Feb 00 - Separate routine to refresh authorisation codes combo
' Reads role codes from Security DB
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsRoles As ADODB.Recordset
    
    cboAuthorisationLevel.Clear
    cboAuthorisationLevel.Tag = ""
    'SDM 14/02/00 SR3000
    ' Display Role codes rather than the descriptions
    ' Always add a blank item at the beginning
    cboAuthorisationLevel.AddItem ""
    sSQL = "Select RoleCode From Role "
    Set rsRoles = New ADODB.Recordset
    rsRoles.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not rsRoles.EOF Then
        rsRoles.MoveFirst
        Do While Not rsRoles.EOF
            cboAuthorisationLevel.AddItem rsRoles!RoleCode
            rsRoles.MoveNext
        Loop
    End If
    rsRoles.Close
    Set rsRoles = Nothing
    
    ' NCJ 11/2/00 - Initialise authorisation code warning
    mbAuthCodeWarningGiven = False

End Sub

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsRoles As ADODB.Recordset

    On Error GoTo ErrHandler
    
    mlGridWidth = Me.ScaleWidth - 120
    
    mbFormLoading = True
    'TA 13/12/2000: offset changed to 600 though this needs tidying up
    Me.Top = frmMenu.tlbMenu.Height + 600
    Me.Height = frmMenu.Height - frmMenu.tlbMenu.Height * 3.25 - frmMenu.sbrMenu.Height - 200
    Me.Left = 0
    Me.Width = frmMenu.Width / 4
    
    picCRFPageDefinition.Height = Picture1.Top - picCRFPageDefinition.Top
    
    'call chkMore_Click to make additional sizing changes
    chkMore_Click
    
    ' Turn on key preview for form, so that F1 (Help) can be trapped by form
    Me.KeyPreview = True
        
    ' NCJ 22 Feb 00 - Code moved to separate routine
    Call RefreshAuthorisationCodes
    
    '   Set flexgrid columns
    ' PN change 20
    With flexCRFElement
        .Cols = mCOL_COUNT
        
        .Row = 0
        .Col = mCOL_FIELD_ORDER
        .ColWidth(mCOL_FIELD_ORDER) = 960
        .Text = "Field order"
        
        ' NCJ 13/12/01 - Added QGroupFieldOrder and Item Type
        .Col = mCOL_QGFIELD_ORDER
        .ColWidth(mCOL_QGFIELD_ORDER) = 960
        .Text = "Group order"
        .Col = mCOL_ITEM_TYPE
        .ColWidth(mCOL_ITEM_TYPE) = 1250
        .Text = "Item type"
        
        .Col = mCOL_DATA_ITEM_NAME
        .ColWidth(mCOL_DATA_ITEM_NAME) = 1450
        .Text = "Name"
        .Col = mCOL_CODE
        .ColWidth(mCOL_CODE) = 1450
        .Text = "Code"
        .Col = mCOL_OPTIONAL
        .ColWidth(mCOL_OPTIONAL) = 800
        .CellAlignment = flexAlignCenterCenter
        .Text = "Optional"
        .Col = mCOL_HIDDEN
        .ColWidth(mCOL_HIDDEN) = 800
        .CellAlignment = flexAlignCenterCenter
        .Text = "Hidden"
        .Col = mCOL_LOCAL
        .ColWidth(mCOL_LOCAL) = 800
        .CellAlignment = flexAlignCenterCenter
        .Text = "Local"
        .Col = mCOL_ONLY_COLLECT_IF
        .ColWidth(mCOL_ONLY_COLLECT_IF) = 2900
        .CellAlignment = flexAlignCenterCenter
        .Text = "Only collect data if"
        
        ' PN
        .Col = mCOL_MANDATORY
        .CellAlignment = flexAlignCenterCenter
        .ColWidth(mCOL_MANDATORY) = 930
        .Text = "Mandatory"
        .Col = mCOL_REASON_FOR_CHANGE
        .CellAlignment = flexAlignCenterCenter
        .ColWidth(mCOL_REASON_FOR_CHANGE) = 800
        .Text = "RFC"
        
        .ColWidth(mCOL_CRFELEMENT_ID) = 0
        .ColWidth(mCOL_QGROUPID) = 0        ' NCJ 14 May 03
        .ColWidth(mCOL_ELEMENTUSE) = 0      ' NCJ 15 May 03
        
        ' PN change - 03/09/99
        ' new requires authorisation field
        .Col = mCOL_AUTHORISATION_LEVEL
        .CellAlignment = flexAlignCenterCenter
        .ColWidth(mCOL_AUTHORISATION_LEVEL) = GetMaxTextWidth
        .Text = "Requires Authorisation"

        'TA 14/09/2000: changes for lab test date expression
        .Col = mCOL_LABTESTDATE
        .ColWidth(mCOL_LABTESTDATE) = 2700
        .CellAlignment = flexAlignCenterCenter
        .Text = "Laboratory Test Date"
        
        ' NCJ 13/12/01 - Show status flag
        'ASH 14/1/2003 - changed "Status Flag" to "Show Icons"
        .Col = mCOL_SHOWFLAG
        .CellAlignment = flexAlignCenterCenter
        .ColWidth(mCOL_SHOWFLAG) = 1000
        .Text = "Show Icons"

        'ZA 13/09/2002 - added Data item length and display length properties
        'ASH 30/10/2002 - changed "Data Item Length" to "Question Length"
        .Col = mCOL_DATA_ITEM_LENGTH
        .CellAlignment = flexAlignCenterCenter
        .ColWidth(mCOL_DATA_ITEM_LENGTH) = 1500
        .Text = "Question Length"

        .Col = mCOL_DISPLAY_LENGTH
        .CellAlignment = flexAlignCenterCenter
        .ColWidth(mCOL_DISPLAY_LENGTH) = 1500
        .Text = "Display Length"
        
        'TA 18/06/2003: new description column
        .Col = mCOL_DESCRIPTION
        .CellAlignment = flexAlignCenterCenter
        .ColWidth(mCOL_DESCRIPTION) = 4500
        .Text = "Description"
        
        .SelectionMode = flexSelectionFree
        .FocusRect = flexFocusHeavy
    End With
    
    ' NCJ 22 Feb 00 - This check box gets set in RefreshFields
'    'added by Mo Morris 19/9/99
'    If gnDisplayNumbers Then
'        chkDisplayNumbers.Value = vbChecked
'    Else
'        chkDisplayNumbers.Value = vbUnchecked
'    End If
    
    txtExpression.Enabled = False
    txtLabTestDateExpr.Enabled = False
    
    'ASH 4/11/2002
    txtDisplayLength.MaxLength = 3
    txtDescription.MaxLength = 255
    
    mbFormLoading = False
    


Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "Form_Load", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If
    
End Sub

'---------------------------------------------------------------------
Private Function GetMaxTextWidth() As Long
'---------------------------------------------------------------------
Dim iIndex As Integer
Dim lMaxWidth As Long
Dim bFontBold  As Boolean
Dim bFontItalic As Boolean
Dim iFontSize As Integer
Dim sFontName As String
    
    On Error GoTo ErrHandler
    
    ' cache the font properties
    bFontBold = Me.FontBold
    bFontItalic = Me.FontItalic
    iFontSize = Me.FontSize
    sFontName = Me.FontName
    
    ' set the font properties of the combo ontot the form
    Me.FontBold = cboAuthorisationLevel.FontBold
    Me.FontItalic = cboAuthorisationLevel.FontItalic
    Me.FontSize = cboAuthorisationLevel.FontSize
    Me.FontName = cboAuthorisationLevel.FontName
    
    For iIndex = 0 To cboAuthorisationLevel.ListCount - 1
        If lMaxWidth < GetTextWidth(cboAuthorisationLevel.List(iIndex)) Then
            lMaxWidth = GetTextWidth(cboAuthorisationLevel.List(iIndex))
        End If
    Next iIndex
    
    ' reset the font properties
    Me.FontBold = bFontBold
    Me.FontItalic = bFontItalic
    Me.FontSize = iFontSize
    Me.FontName = sFontName
    GetMaxTextWidth = lMaxWidth
        
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCRFPageDefinition.GetMaxTextWidth"
    
End Function

'---------------------------------------------------------------------
Private Function GetTextWidth(sText As String) As Long
'---------------------------------------------------------------------
' obtain the width of the text from the form control
' this works since the text box and form hasve the same font properties
'---------------------------------------------------------------------
' PN 02/09/99
Dim sDefaultText As String
    
    'WillC 28/2/2000 SR 2897 if nothing is chosen then give the col enough width to show the caption.
    sDefaultText = "Requires Authorisation"
    If sText = "" Then
        sText = sDefaultText
    End If
    
    GetTextWidth = TextWidth(String(Len(sText), "_"))

End Function

'---------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'---------------------------------------------------------------------
' NCJ 1 Jul 03 - Check they mean to close without saving
' NCJ 4 Sept 03 - Don't do this because it's irritating!
'---------------------------------------------------------------------

'    If UnloadMode = vbFormControlMenu Then
'        If DialogQuestion("Are you sure you wish to close? Unsaved data will be lost") = vbNo Then
'            Cancel = 1
'        End If
'    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub Form_Resize()
 '---------------------------------------------------------------------
'resize form
'---------------------------------------------------------------------

    On Error Resume Next
    
    'TA 15/11/2000: only enable more button when not maximised
    chkMore.Enabled = (Me.WindowState = vbNormal)
    
    If Me.WindowState <> vbMinimized Then
        
        If Me.Width >= frmMenu.Width / 4 Then
            flexCRFElement.Width = Me.ScaleWidth - 120
        Else
            Me.Width = frmMenu.Width / 4
        End If
            
        If Me.Height >= 4500 Then
            picCRFPageDefinition.Height = Me.ScaleHeight - Picture1.Height
            flexCRFElement.Height = picCRFPageDefinition.Height - flexCRFElement.Top - cmdMoveUp.Height - 240
            cmdMoveUp.Top = picCRFPageDefinition.Top + flexCRFElement.Top + flexCRFElement.Height + 120
            cmdMoveDown.Top = cmdMoveUp.Top
            cmdRenumber.Top = cmdMoveUp.Top
        Else
            Me.Height = 4500
        End If
        
    End If

End Sub

'---------------------------------------------------------------------
Private Sub txtCRFPageTitle_LostFocus()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    ' Forget it if we're cancelling anyway
    If Me.ActiveControl.Name = "cmdCancel" Then Exit Sub
        
    If Me.Visible = True Then
        If Not gblnValidString(txtCRFPageTitle.Text, valOnlySingleQuotes) Then
            DialogError "eForm title" & gsCANNOT_CONTAIN_INVALID_CHARS
            txtCRFPageTitle.SetFocus
        End If
    End If
    
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "txtCRFPageTitle_LostFocus", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------
    
    If KeyCode = vbKeyF1 Then               ' Show user guide
        'ShowDocument Me.hwnd, gsMACROUserGuidePath
        
        'REM 07/12/01 - New Call for the MACRO Help
        Call MACROHelp(Me.hWnd, App.Title)
        
    End If
    
End Sub

'--------------------------------------------------------------------
Private Sub txtDescription_KeyPress(KeyAscii As Integer)
'--------------------------------------------------------------------
' Allow Return key to terminate input
'--------------------------------------------------------------------
    
    If KeyAscii = Asc(vbCr) Then
        SendKeys vbTab      ' Tab to next field
    End If

End Sub

'---------------------------------------------------------------------
Private Sub txtDescription_LostFocus()
'---------------------------------------------------------------------
' - validate the value entered by user and hide the text box
'---------------------------------------------------------------------
Dim nRow As Integer
Dim sDescription As String
Dim bIsValid As Boolean

    On Error GoTo ErrLabel
    
    ' Forget it if we're cancelling anyway
    If Me.ActiveControl.Name = "cmdCancel" Then Exit Sub
    
    bIsValid = False
    
    sDescription = Trim(txtDescription.Text)
    'validate the data entered by user
    'if a user deletes an existing value, put the dataitemlength value there
    If sDescription = "" Then
        bIsValid = True
    Else
        If Not gblnValidString(sDescription, valOnlySingleQuotes) Then
            'check for invalid characters
            bIsValid = False
            Call DialogWarning("Description" & gsCANNOT_CONTAIN_INVALID_CHARS)
        Else
            bIsValid = True
        End If
    End If
    
    If bIsValid Then
        'retrieve the row where the text box is
        nRow = ExtractFirstItemFromList(txtDescription.Tag, ",")
        flexCRFElement.TextMatrix(nRow, mCOL_DESCRIPTION) = sDescription
        'clear and hide the text box
        txtDescription.Text = ""
        txtDescription.Visible = False
    Else
        If txtDescription.Visible Then
            txtDescription.SetFocus
        End If
    End If
    
    Exit Sub
    
ErrLabel:
     Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, "txtDescription_LostFocus", Err.Source)
                               
        Case OnErrorAction.Retry
            Resume
    End Select
End Sub

'---------------------------------------------------------------------
Private Sub txtDisplayLength_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
' Let them press Return to move on
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler

    If KeyAscii = Asc(vbCr) Then
        SendKeys vbTab      ' Tab to next field
    End If
    
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "txtDisplayLength_KeyPress", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If

End Sub

'---------------------------------------------------------------------
Private Sub txtDisplayLength_LostFocus()
'---------------------------------------------------------------------
' ZA 13/09/2002 - validate the value entered by user and hide the text box
'---------------------------------------------------------------------
Dim nRow As Integer
Dim nDisplayLength As Integer
Dim sDisplayValue As String
Dim bIsValid As Boolean

    On Error GoTo ErrLabel
    
    ' Forget it if we're cancelling anyway
    If Me.ActiveControl.Name = "cmdCancel" Then Exit Sub
    
    bIsValid = False
    
    sDisplayValue = Trim(txtDisplayLength.Text)
    'validate the data entered by user
    'if a user deletes an existing value, put the dataitemlength value there
    If sDisplayValue = "" Then
        'retrieve the row where the text box is
        nRow = ExtractFirstItemFromList(txtDisplayLength.Tag, ",")
        
      '  nDisplayLength = flexCRFElement.TextMatrix(nRow, mCOL_DATA_ITEM_LENGTH)
        nDisplayLength = NULL_INTEGER
        bIsValid = True
    Else
        nDisplayLength = Val(sDisplayValue)
        
        If (nDisplayLength < 1 Or nDisplayLength > 255) Or Not gblnValidString(sDisplayValue, valNumeric) Then
            Call DialogWarning("Display length for a question must be a whole number between 1 and 255, or blank")
        Else
            bIsValid = True
        End If
    End If
    
    If bIsValid Then
        
        'retrieve the row where the text box is
        nRow = ExtractFirstItemFromList(txtDisplayLength.Tag, ",")
                
            flexCRFElement.TextMatrix(nRow, mCOL_DISPLAY_LENGTH) = txtDisplayLength.Text
        
        'clear and hide th text box
        txtDisplayLength.Text = ""
        txtDisplayLength.Visible = False
    Else
        If txtDisplayLength.Visible Then
            txtDisplayLength.SetFocus
        End If
    End If
    
    Exit Sub
    
ErrLabel:
     Select Case MACROErrorHandler(Me.Name, Err.Number, Err.Description, "txtDisplayLength_LostFocus", Err.Source)
                               
        Case OnErrorAction.Retry
            Resume
    End Select
End Sub

'---------------------------------------------------------------------
Private Sub txtFormDate_LostFocus()
'---------------------------------------------------------------------
Dim sFormDate As String

    On Error GoTo ErrHandler

    sFormDate = Trim(txtFormDate.Text)
    
    If Not gblnValidString(sFormDate, valOnlySingleQuotes) Then
        MsgBox "Form date" & gsCANNOT_CONTAIN_INVALID_CHARS, _
                vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
        txtFormDate.SetFocus
    ' Use Arezzo expression checker - NCJ 13/8/99
    'Mo Morris 16/9/99 Null string check added before call to IsValidExpression
    ElseIf sFormDate <> vbNullString Then
        If Not gclmGuideline.IsValidExpression(sFormDate) Then
            MsgBox "Form date is not a valid expression.", _
                vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
            txtFormDate.SetFocus
        End If
    End If
    
    
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "txtFormDate_LostFocus", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub txtExpression_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    If KeyAscii = Asc(vbCr) Then
        If cmdValidationExpression.Visible Then
            cmdValidationExpression.Visible = False
        End If
        SendKeys vbTab
    End If
    
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "txtExpression_KeyPress", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If
        
End Sub

'---------------------------------------------------------------------
Private Sub txtExpression_LostFocus()
'---------------------------------------------------------------------
' This is an element's "Collect if" condition
' NCJ 14 May 03 - Check length of condition
'---------------------------------------------------------------------
Dim nStartRow As Integer
Dim nEndRow As Integer
Dim sStartEnd As String
Dim nRow As Integer
Dim sExpr As String

    On Error GoTo ErrHandler
    
    If Not Me.Visible Then Exit Sub
    
    ' Forget it if we're cancelling anyway
    If Me.ActiveControl.Name = "cmdCancel" Then Exit Sub
    
    If Me.ActiveControl.Name <> "cmdValidationExpression" And Me.ActiveControl.Name <> "txtExpression" Then
        'skip condition entered
        sExpr = Trim(txtExpression.Text)
        If sExpr > "" Then
            If Not gblnValidString(sExpr, valOnlySingleQuotes) Then
            
                ' PN 28/09/99 when scrolling the skip condition may be made invisible
                If txtExpression.Visible Then
                    DialogError "Condition" & gsCANNOT_CONTAIN_INVALID_CHARS
                    txtExpression.SetFocus
                End If
                Exit Sub
                
            ' NCJ 14 May 03 - Check length of Collect If condition
            ElseIf Len(sExpr) > glMAX_AREZZO_EXPR_LEN Then
                
                ' PN 28/09/99 when scrolling the skip condition may be made invisible
                If txtExpression.Visible Then
                    DialogError "Condition may not be longer than " & glMAX_AREZZO_EXPR_LEN & " characters"
                    txtExpression.SetFocus
                    Exit Sub
                End If
                
            ElseIf Not gclmGuideline.IsValidCondition(sExpr) Then
                
                ' PN 28/09/99 when scrolling the skip condition may be made invisible
                If txtExpression.Visible Then
                    DialogError "Condition is not a valid AREZZO term"
                    txtExpression.SetFocus
                    Exit Sub
                End If
            
            Else
                    sStartEnd = txtExpression.Tag
                    nStartRow = ExtractFirstItemFromList(sStartEnd, ",")
                    nEndRow = sStartEnd
                    For nRow = nStartRow To nEndRow
                        flexCRFElement.TextMatrix(nRow, mCOL_ONLY_COLLECT_IF) = sExpr
                    Next
            End If
        Else
                    sStartEnd = txtExpression.Tag
                    nStartRow = ExtractFirstItemFromList(sStartEnd, ",")
                    nEndRow = sStartEnd
                    For nRow = nStartRow To nEndRow
                        flexCRFElement.TextMatrix(nRow, mCOL_ONLY_COLLECT_IF) = sExpr
                    Next
        End If

        txtExpression.Enabled = False
        txtExpression.Visible = False
        cmdValidationExpression.Visible = False
    
    End If
    
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "txtExpression_LostFocus", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub SaveFormDetails()
'---------------------------------------------------------------------
' NCJ 22 Feb 00 - Combined some SQL calls into one
'---------------------------------------------------------------------
Dim nRow As Integer
Dim nOptional As Integer
Dim nHidden As Integer
Dim nLocal As Integer
Dim nShowFlag As Integer
Dim iMandatory As Byte
Dim iSequentialEntry As Byte
Dim iReasonForChange As Byte
' PN change 23 - 06/09/99
Dim sAuthorisationCode As String
Dim oForm As Form
Dim nElementUse As eElementUse

'ZA 03/09/2002 - Save Display length value for a data item
Dim sDisplayLength As String
Dim nDisplayLength As Integer
Dim sDataItemLength As String

'ASH 4/11/2002 - Value for New field
Dim lEformWidth As Long
Dim sDescription As String

    On Error GoTo ErrHandler
    
    ' NCJ 15 May 06 - Ensure we don't save in RO access
    If meAccessMode = sdReadOnly Then Exit Sub
    
    'ASH 4/11/2002 - Value for New field
    If optPortrait.Value Then
        lEformWidth = glPORTRAIT_WIDTH
    Else
        lEformWidth = glLANDSCAPE_WIDTH
    End If
    
    'SDM 06/01/00 SR2392    Added HideIfInactive
    ' NCJ 28 Jan 03 - Replaced chkPrompt.Value with 0
    UpdateCRFPageDetails Me.ClinicalTrialId, Me.VersionId, mlCRFPageId, _
                    Trim(txtCRFPageTitle.Text), Trim(cboFormLabel.Text), chkLocal.Value, _
                    Trim(txtFormDate.Text), chkSequentialEntry.Value, _
                    chkHideIfInactive.Value, chkDisplayNumbers.Value, 0, lEformWidth
    
    With flexCRFElement
        For nRow = 1 To .Rows - 1
            ' NCJ 15 May 03 - Don't update field order for group questions
            ' because doing it for the group automatically does all its group questions
            If .TextMatrix(nRow, mCOL_QGFIELD_ORDER) = "" Then
                UpdateCRFFieldOrder Me.ClinicalTrialId, Me.VersionId, mlCRFPageId, _
                                    .TextMatrix(nRow, mCOL_CRFELEMENT_ID), _
                                    .TextMatrix(nRow, mCOL_FIELD_ORDER), _
                                    .TextMatrix(nRow, mCOL_QGROUPID)
            End If
            
            ' Now collect the other bits of information about the element
            If .TextMatrix(nRow, mCOL_OPTIONAL) = "" Then
                nOptional = 0
            Else
                nOptional = 1
            End If
            If .TextMatrix(nRow, mCOL_HIDDEN) = "" Then
                nHidden = 0
            Else
                nHidden = 1
            End If
            If .TextMatrix(nRow, mCOL_LOCAL) = "" Then
                nLocal = 0
            Else
                nLocal = 1
            End If
            ' NCJ 13/12/01
            If .TextMatrix(nRow, mCOL_SHOWFLAG) = "" Then
                nShowFlag = 0
            Else
                nShowFlag = 1
            End If
            
            ' PN change 20
            If chkSequentialEntry.Value = 0 Then
                iSequentialEntry = 0
            Else
                iSequentialEntry = 1
            End If
            If .TextMatrix(nRow, mCOL_MANDATORY) = vbNullString Then
                iMandatory = 0
            Else
                iMandatory = 1
            End If
            If .TextMatrix(nRow, mCOL_REASON_FOR_CHANGE) = vbNullString Then
                iReasonForChange = 0
            Else
                iReasonForChange = 1
            End If
            
            'check if this question is the form date question
'            If .TextMatrix(nRow, mCOL_CODE) = cboFormDate.Text Then
'                nElementUse = eElementUse.EFormVisitDate
'            Else
'                nElementUse = eElementUse.User
'            End If
            
            ' PN change 23 - 06/09/99
'            sAuthorisationCode = moAuthorisationLevels.Key(.TextMatrix(nRow, mCOL_AUTHORISATION_LEVEL))
            ' NCJ 15/2/00 - The role code is stored in the text matrix
            sAuthorisationCode = .TextMatrix(nRow, mCOL_AUTHORISATION_LEVEL)
            ' NCJ 15/2/00 - Use special code of "0" for no role code
            If sAuthorisationCode = "" Then
                sAuthorisationCode = "0"
            End If
            
            'ZA 13/09/2002 - add Display length value
            'only for those questions where data item length is shown
            sDataItemLength = Trim(.TextMatrix(nRow, mCOL_DATA_ITEM_LENGTH))
            nDisplayLength = NULL_INTEGER
            If sDataItemLength <> "" Then
                sDisplayLength = Trim(.TextMatrix(nRow, mCOL_DISPLAY_LENGTH))
                If sDisplayLength <> "" Then
                    nDisplayLength = Val(sDisplayLength)
                End If
            End If
            
            'TA 18/06/2003: get description
             sDescription = Trim(.TextMatrix(nRow, mCOL_DESCRIPTION))
            
            ' NCJ 13/12/01 - Added nShowFlag
            ' NCJ 15 May 03 - Pass in ElementUse directly from TextMatrix
            UpdateCRFElementProperties Me.ClinicalTrialId, Me.VersionId, _
                    mlCRFPageId, .TextMatrix(nRow, mCOL_CRFELEMENT_ID), nOptional, _
                    nHidden, nLocal, .TextMatrix(nRow, mCOL_ONLY_COLLECT_IF), _
                    iMandatory, iReasonForChange, sAuthorisationCode, .TextMatrix(nRow, mCOL_LABTESTDATE), _
                    nShowFlag, .TextMatrix(nRow, mCOL_ELEMENTUSE), nDisplayLength, sDescription
            
        Next
    End With
    
    ' PN change 23 - 06/09/99
    Set oForm = frmMenu.gFindForm(ClinicalTrialId, VersionId, "frmCRFDesign")
    If Not oForm Is Nothing Then
        frmCRFDesign.tabCRF.Tabs(gsCRF_PAGE_LABEL & CStr(Me.CRFPageId)).Caption = Replace(Trim(txtCRFPageTitle.Text), "&", "&&")
        ' NCJ 22 Feb 00 - Only update if CRFPageId matches
        If oForm.CRFPageId = Me.CRFPageId Then
            Call frmCRFDesign.RefreshMe
        End If
    End If
    
    ' NCJ 24/11/00 - Use RefreshQuestionLists instead of calling frmDataList
    ' NCJ 15 Jan 02 - Use frmMenu's routine
    Call frmMenu.RefreshQuestionLists(Me.ClinicalTrialId)

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmCRFPageDefinition.SaveFormDetails"
    
End Sub

'---------------------------------------------------------------------
Private Sub txtLabTestDateExpr_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    If KeyAscii = Asc(vbCr) Then
        If cmdValidationExpression.Visible Then
            cmdValidationExpression.Visible = False
        End If
        SendKeys vbTab
    End If
    
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "txtLabTestDateExpr_KeyPress", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If

End Sub

'---------------------------------------------------------------------
Private Sub txtLabTestDateExpr_LostFocus()
'---------------------------------------------------------------------
' TA 27/09/2000: check lab test date expression
' NCJ 14 May 03 - Check length of expression
'---------------------------------------------------------------------
Dim nRow As Long
Dim sExpr As String

    On Error GoTo ErrHandler
    
    If Not Me.Visible Then Exit Sub
    
    ' Forget it if we're cancelling anyway
    If Me.ActiveControl.Name = "cmdCancel" Then Exit Sub
    
    'get row that was originally selected
    nRow = ExtractFirstItemFromList(txtLabTestDateExpr.Tag, ",")
    
    If Me.ActiveControl.Name <> "cmdValidationExpression" And Me.ActiveControl.Name <> "txtLabTestDateExpr" Then

        sExpr = Trim(txtLabTestDateExpr.Text)
        'TA 14/09/2000: lab test date expression checks
        If sExpr > "" Then
            If Not gblnValidString(sExpr, valOnlySingleQuotes) Then
            
                ' PN 28/09/99 when scrolling the lab test date expr may be made invisible
                If txtLabTestDateExpr.Visible Then
                    DialogError "Expression" & gsCANNOT_CONTAIN_INVALID_CHARS
                    txtLabTestDateExpr.SetFocus
                End If
                Exit Sub
                
            ' NCJ 14 May 03 - Check length of Collect If condition
            ElseIf Len(sExpr) > glMAX_AREZZO_EXPR_LEN Then
                
                ' PN 28/09/99 when scrolling the lab test date expr may be made invisible
                If txtLabTestDateExpr.Visible Then
                    DialogError "Expression may not be longer than " & glMAX_AREZZO_EXPR_LEN & " characters"
                    txtLabTestDateExpr.SetFocus
                    Exit Sub
                End If
                
            ElseIf Not gclmGuideline.IsValidExpression(sExpr) Then
                
                ' PN 28/09/99 when scrolling the lab test date expr may be made invisible
                If txtLabTestDateExpr.Visible Then
                    DialogError "Expression is not a valid AREZZO term"
                    txtLabTestDateExpr.SetFocus
                    Exit Sub
                End If
            Else
                'store data
                flexCRFElement.TextMatrix(nRow, mCOL_LABTESTDATE) = sExpr
            End If
        Else
            'store data
            flexCRFElement.TextMatrix(nRow, mCOL_LABTESTDATE) = sExpr
        End If
    
        txtLabTestDateExpr.Enabled = False
        txtLabTestDateExpr.Visible = False
        cmdValidationExpression.Visible = False
            
    End If
        
Exit Sub
ErrHandler:
    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "txtLabTestDateExpr_LostFocus", Err.Source) = OnErrorAction.Retry Then
            Resume
    End If
End Sub

'---------------------------------------------------------------------
Private Sub EnableMoveUpMoveDown(ByVal nRow As Integer)
'---------------------------------------------------------------------
' NCJ 15 May 03
' Enable or disable the move up/down buttons according to row clicked in
' NCJ 30 Jun 03 - Bug 1900 - Reversed the meanings of "up" and "down" to match 2.2
' NCJ 15 May 06 - Consider Access Mode
'---------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    ' Default to disabled
    cmdMoveDown.Enabled = False
    cmdMoveUp.Enabled = False
    
    ' NCJ 15 May 06 - If we're read-only, forget it
    If meAccessMode = sdReadOnly Then Exit Sub
    
    ' If not in a row then forget it
    If nRow < 1 Then Exit Sub
    
    ' If it's a group question we don't allow move up/down
    If flexCRFElement.TextMatrix(nRow, mCOL_QGFIELD_ORDER) > "" Then Exit Sub

    ' If it's a hidden question we don't allow move up/down
    If flexCRFElement.TextMatrix(nRow, mCOL_HIDDEN) = msSELECTEDTEXT Then Exit Sub

    ' If it has field order 0 we don't allow move up/down (e.g. eForm date qu)
    If CInt(flexCRFElement.TextMatrix(nRow, mCOL_FIELD_ORDER)) = 0 Then Exit Sub

    ' If it's the "bottom" element we don't allow move down
    If CInt(flexCRFElement.TextMatrix(nRow, mCOL_FIELD_ORDER)) = mnMaxFieldOrder Then
        If CInt(flexCRFElement.TextMatrix(nRow, mCOL_FIELD_ORDER)) > mnMinFieldOrder Then
            ' It's not already at the top so allow move up
            cmdMoveUp.Enabled = True
        End If
        Exit Sub
    End If
    
    ' If it's the "top" element we don't allow move up
    If CInt(flexCRFElement.TextMatrix(nRow, mCOL_FIELD_ORDER)) = mnMinFieldOrder Then
        If CInt(flexCRFElement.TextMatrix(nRow, mCOL_FIELD_ORDER)) < mnMaxFieldOrder Then
            ' It's not already at the bottom so allow move down
            cmdMoveDown.Enabled = True
        End If
        Exit Sub
    End If

    ' If we get here it's OK!
    cmdMoveDown.Enabled = True
    cmdMoveUp.Enabled = True

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmCRFPageDefinition.EnableMoveUpMoveDown"

End Sub

'---------------------------------------------------------------------
Private Sub SwapElements(nStartRow1 As Integer, nEndRow1 As Integer, nEndRow2 As Integer)
'---------------------------------------------------------------------
' NCJ 15 May 03
' Swap two adjacent blocks of rows in the grid
' by copying the cell contents into a collection and then putting them back in the right order
' nStartRow1 to nEndRow1 is the first block of rows and the second block finishes at nEndRow2
' Assume the second block starts where first one finishes
' May have nStartRow1 = nEndRow1 and/or nStartRow2 = nEndRow2
'---------------------------------------------------------------------
Dim colRowStore As Collection
Dim nRow As Integer
Dim nCol As Integer
Dim i As Integer
Dim sFieldOrder1 As String
Dim sFieldOrder2 As String
Dim nStartRow2 As Integer

    On Error GoTo ErrLabel
    
    Set colRowStore = New Collection
    
    With flexCRFElement
    
        ' Assume second block starts where first one finishes
        nStartRow2 = nEndRow1 + 1
        
        ' Find what field orders we're swapping
        sFieldOrder1 = .TextMatrix(nStartRow1, mCOL_FIELD_ORDER)
        sFieldOrder2 = .TextMatrix(nStartRow2, mCOL_FIELD_ORDER)
        
        ' Store the grid cell values in the collection in the right order,
        ' starting with the second block, giving them FieldOrder1
        For nRow = nStartRow2 To nEndRow2
            ' Column 0 is the field order
            colRowStore.Add sFieldOrder1
            For nCol = 1 To mCOL_COUNT - 1
                colRowStore.Add .TextMatrix(nRow, nCol)
            Next nCol
        Next nRow
        ' Now add the data from the first block, giving them FieldOrder2
        For nRow = nStartRow1 To nEndRow1
            ' Column 0 is the field order
            colRowStore.Add sFieldOrder2
            For nCol = 1 To mCOL_COUNT - 1
                colRowStore.Add .TextMatrix(nRow, nCol)
            Next nCol
        Next nRow
        
        ' Now shove the data back in from the beginning
        ' (with block 1 and block 2 reversed)
        i = 1
        For nRow = nStartRow1 To nEndRow2
            For nCol = 0 To mCOL_COUNT - 1
                .TextMatrix(nRow, nCol) = colRowStore.Item(i)
                i = i + 1
            Next nCol
        Next nRow
        
    End With

    Set colRowStore = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmCRFPageDefinition.SwapElements"

End Sub

'---------------------------------------------------------------------
Private Sub SetFieldsAccess()
'---------------------------------------------------------------------
' NCJ 15 May 06 - Set access to editing fields according to access mode
'---------------------------------------------------------------------

    If meAccessMode = sdReadOnly Then
        txtCRFPageTitle.Enabled = False
        cboFormDate.Enabled = False
        cboFormLabel.Enabled = False
        cmdFormLabelExpression.Visible = False
        chkHideIfInactive.Enabled = False
        chkDisplayNumbers.Enabled = False
        optPortrait.Enabled = False
        optLandscape.Enabled = False
        cmdRenumber.Enabled = False
        cmdApply.Enabled = False
    End If
    
End Sub
