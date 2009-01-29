VERSION 5.00
Begin VB.Form frmQueryOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MACRO Query Output Options"
   ClientHeight    =   10125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10125
   ScaleWidth      =   7515
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Caption         =   "Output File Name:"
      Height          =   3050
      Left            =   100
      TabIndex        =   38
      Top             =   6500
      Width           =   7300
      Begin VB.OptionButton optNoStamp 
         Caption         =   "No Date or Time Stamp"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   2700
         Width           =   2055
      End
      Begin VB.CheckBox chkUseStudyName 
         Caption         =   "Save using Study Name"
         Height          =   255
         Left            =   100
         TabIndex        =   45
         Top             =   700
         Width           =   3855
      End
      Begin VB.CheckBox chkSaveInAppPath 
         Caption         =   "Save in MACRO Application Path\Out Folder"
         Height          =   255
         Left            =   100
         TabIndex        =   44
         Top             =   300
         Width           =   3855
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   4200
         TabIndex        =   43
         Top             =   500
         Width           =   3000
      End
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   4200
         TabIndex        =   42
         Top             =   950
         Width           =   3000
      End
      Begin VB.TextBox txtUserSpecName 
         Height          =   315
         Left            =   100
         TabIndex        =   41
         Top             =   1300
         Width           =   3800
      End
      Begin VB.OptionButton optDateTimeStamp 
         Caption         =   "Date/Time Stamp (yyyymmddhhmmss)"
         Height          =   255
         Left            =   100
         TabIndex        =   40
         Top             =   2400
         Width           =   3015
      End
      Begin VB.OptionButton optDateStamp 
         Caption         =   "Date Stamp (yyyymmdd)"
         Height          =   255
         Left            =   100
         TabIndex        =   39
         Top             =   2100
         Width           =   2055
      End
      Begin VB.Label Label8 
         Caption         =   "Specify File Name"
         Height          =   315
         Left            =   105
         TabIndex        =   49
         Top             =   1050
         Width           =   1500
      End
      Begin VB.Label Label7 
         Caption         =   "Specify Path"
         Height          =   315
         Left            =   4200
         TabIndex        =   48
         Top             =   255
         Width           =   1200
      End
      Begin VB.Label Label6 
         Caption         =   "File Name to include: -"
         Height          =   255
         Left            =   100
         TabIndex        =   46
         Top             =   1800
         Width           =   1695
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Output Subject Label Option"
      Height          =   800
      Left            =   120
      TabIndex        =   33
      Top             =   2700
      Width           =   4400
      Begin VB.CheckBox chkExcludeLabel 
         Caption         =   "Exclude Subject Label from Saved Output Files"
         Height          =   435
         Left            =   120
         TabIndex        =   34
         Top             =   300
         Width           =   3225
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Question Code Options"
      Height          =   1400
      Left            =   100
      TabIndex        =   31
      Top             =   3700
      Width           =   4400
      Begin VB.TextBox txtShortCodeLength 
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   950
         Width           =   400
      End
      Begin VB.OptionButton optLongCodes 
         Caption         =   "Long Codes - VisitCode/eFormCode/QuestionCode"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   4035
      End
      Begin VB.OptionButton optShortCodes 
         Caption         =   "Short Codes - SAS style codes"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label Label5 
         Caption         =   "Short Code length (valid between 8 and 18)"
         Height          =   255
         Left            =   650
         TabIndex        =   36
         Top             =   1000
         Width           =   3500
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Screen Output Options"
      Height          =   2400
      Left            =   120
      TabIndex        =   29
      Top             =   100
      Width           =   4400
      Begin VB.CheckBox chkRepeatNumber 
         Caption         =   "Repeat Number"
         Height          =   255
         Left            =   2300
         TabIndex        =   32
         Top             =   1200
         Width           =   2000
      End
      Begin VB.CheckBox chkPersonId 
         Caption         =   "PersonId"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1500
         Width           =   2000
      End
      Begin VB.CheckBox chkSplitGrid 
         Caption         =   "Split bar between identification fields and response fields"
         Height          =   435
         Left            =   120
         TabIndex        =   7
         Top             =   1900
         Width           =   3225
      End
      Begin VB.CheckBox chkFormCycle 
         Caption         =   "Form cycle number"
         Height          =   255
         Left            =   2300
         TabIndex        =   6
         Top             =   900
         Width           =   2000
      End
      Begin VB.CheckBox chkVisitCycle 
         Caption         =   "Visit cycle number"
         Height          =   255
         Left            =   2300
         TabIndex        =   5
         Top             =   600
         Width           =   2000
      End
      Begin VB.CheckBox chkLabel 
         Caption         =   "Subject Label"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   2000
      End
      Begin VB.CheckBox chkSiteCode 
         Caption         =   "Site Code"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   900
         Width           =   2000
      End
      Begin VB.CheckBox chkStudyName 
         Caption         =   "Study Name"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   2000
      End
      Begin VB.Line Line1 
         X1              =   100
         X2              =   4200
         Y1              =   1850
         Y2              =   1850
      End
      Begin VB.Label Label4 
         Caption         =   "Select the  identification fields that you wish to display"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   300
         Width           =   4095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Output Category Codes or Values"
      Height          =   1000
      Left            =   100
      TabIndex        =   28
      Top             =   5300
      Width           =   4400
      Begin VB.OptionButton optValues 
         Caption         =   "Category responses to be Values (e.g. Female, Male)"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   4095
      End
      Begin VB.OptionButton optCodes 
         Caption         =   "Category responses to be Codes (e.g. F, M)"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   300
         Width           =   3555
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6200
      TabIndex        =   23
      Top             =   9650
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4900
      TabIndex        =   22
      Top             =   9650
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      Caption         =   "Missing Response Special Values"
      Height          =   2400
      Left            =   4600
      TabIndex        =   24
      Top             =   100
      Width           =   2800
      Begin VB.TextBox txtNotApplicable 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   1950
         Width           =   2600
      End
      Begin VB.TextBox txtUnobtainable 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   1250
         Width           =   2600
      End
      Begin VB.TextBox txtMissing 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   550
         Width           =   2600
      End
      Begin VB.Label Label3 
         Caption         =   "For Status NotApplicable insert"
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   1700
         Width           =   2205
      End
      Begin VB.Label Label2 
         Caption         =   "For Status Unobtainable insert"
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   1000
         Width           =   2205
      End
      Begin VB.Label Label1 
         Caption         =   "For Status Missing insert"
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   300
         Width           =   2205
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Output To:"
      Height          =   3600
      Left            =   4600
      TabIndex        =   0
      Top             =   2700
      Width           =   2800
      Begin VB.CheckBox chkSASInformatColons 
         Caption         =   "Precede SAS informats with colons. (e.g. :$11., :5.)"
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   3100
         Width           =   2235
      End
      Begin VB.OptionButton optSTATA 
         Caption         =   "STATA (Float dates)"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Uses ddmmyyyy Float dates (e.g. 01012004 for 1 January 2004)"
         Top             =   2100
         Width           =   2300
      End
      Begin VB.OptionButton optMacroBD 
         Caption         =   "MACRO Batch Data Format"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2550
         Width           =   2300
      End
      Begin VB.OptionButton optSTATAStandardDates 
         Caption         =   "STATA (Standard dates)"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Uses ddmmmyyy Standard dates (e.g. 01jan2004)"
         Top             =   1650
         Width           =   2300
      End
      Begin VB.OptionButton optSAS 
         Caption         =   "SAS"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   2300
      End
      Begin VB.OptionButton optSPSS 
         Caption         =   "SPSS"
         Height          =   195
         Left            =   1920
         TabIndex        =   17
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton optAccess 
         Caption         =   "Access"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   750
         Width           =   2300
      End
      Begin VB.OptionButton optCSV 
         Caption         =   "Comma Separated Values"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   300
         Width           =   2300
      End
      Begin VB.Line Line2 
         X1              =   100
         X2              =   2600
         Y1              =   2950
         Y2              =   2950
      End
   End
End
Attribute VB_Name = "frmQueryOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
' File:         frmQueryOptions
' Copyright:    InferMed Ltd. 2000. All Rights Reserved
' Author:       Mo Morris, February 2002
' Purpose:      Contains conttols for setting/changing the Query Output Options
'----------------------------------------------------------------------------------------'
'   Revisions:
'
'   Mo  7/5/2002    changes stemming from Label/PersonId being switched
'                   from a single field to 2 separate fields.
'                   chkLabelPersonId replaced by chkLabel and chkPersonId.
'                   chkSubjectLabels removed.
'   Mo  12/11/2002  facilities for changing between Long and Short
'                   Question codes added. Module level variables correctly renamed
'   Mo  27/1/2003   Changes throughout Query module as RQGs are incorporated.
'   Mo  15/7/2003   Minor change, optSPSS.visible set to false until SPSS output has been written
'   Mo  17/11/2004  Bug 2411 - There are now 2 forms of STATA output Standard and Float.
'   Mo  17/11/2004  Bug 2411 - OutputToSTATA now works in 2 ways:-
'                   "Standard"  Uses ddmmmyyyy Standard dates (e.g. 01jan2004)
'                   "Float"     Uses ddmmyyyy Float dates (e.g. 01012004 for 1 January 2004)
'                   New sub optSTATAStandardDates_Click has been added.
'                   Changes have been made to Form_Load, ActivateChanges
'   Mo  30/5/2006   Bug 2668 - Option to exclude Subject Label from saved output files
'   Mo  2/6/2006    Bug 2737 - Add Question Short Code length to the Options Window
'   Mo  9/6/2006    Bug 2739 - Changes to STATA export of Special Values on strings
'                   No code changes, just changes to position of controls.
'   Mo  1/11/2006   Bug 2795 - "Precede SAS informats with colons" CheckBox added to the options window.
'                   AssessSVandBD renamed AssessOptions and now enables/disables chkSASInformatColons.
'   Mo  2/4/2007    MRC15022007 - Query Module Batch Facilities
'                   New controls Frame7, chkSaveInAppPath, chkUseStudyName, txtUserSpecName, Label6
'                   optDateStamp, optDateTimeStamp, optNoStamp Drive1, Dir1 added
'----------------------------------------------------------------------------------------'

Option Explicit

Private mbSVMissingChanged As Boolean
Private mbSVUnobtainableChanged As Boolean
Private mbSVNotApplicableChanged As Boolean
Private mbOutPutTypeChanged As Boolean
Private mbOutputCodeValuesChanged As Boolean
Private mbDisplayStudyNameChanged As Boolean
Private mbDisplaySiteCodeChanged As Boolean
Private mbDisplayLabelChanged As Boolean
Private mbDisplayPersonIdChanged As Boolean
Private mbDisplayVisitCycleChanged As Boolean
Private mbDisplayFormCycleChanged As Boolean
Private mbDisplayRepeatNumberChanged As Boolean
Private mbSplitGridChanged As Boolean
Private mbLongShortQuestionsChanged As Boolean
'Mo 30/5/2006 Bug 2668
Private mbExcludeLabelChanged As Boolean
'Mo 1/11/2006 Bug 2795
Private mbSASInformatColonsChanged As Boolean
'Mo 2/4/2007 MRC15022007
Private mbFileNamePathChanged As Boolean
Private mbFileNameTextChanged As Boolean
Private mbFileNameStampChanged As Boolean

'--------------------------------------------------------------------
Private Sub chkExcludeLabel_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    mbExcludeLabelChanged = True

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "chkExcludeLabel_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub chkFormCycle_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    mbDisplayFormCycleChanged = True

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "chkFormCycle_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub chkLabel_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler

    mbDisplayLabelChanged = True

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "chkLabel_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub chkPersonId_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler

    mbDisplayPersonIdChanged = True

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "chkPersonId_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub chkRepeatNumber_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    mbDisplayRepeatNumberChanged = True

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "chkRepeatNumber_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub chkSASInformatColons_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    mbSASInformatColonsChanged = True
    mbOutPutTypeChanged = True

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "chkSASInformatColons_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'Mo 2/4/2007 MRC15022007
'--------------------------------------------------------------------
Private Sub chkSaveInAppPath_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    If chkSaveInAppPath.Value = 1 Then
        Drive1.Drive = Mid(gsOUT_FOLDER_LOCATION, 1, InStr(gsOUT_FOLDER_LOCATION, "\"))
        Drive1.Enabled = False
        Dir1.Path = gsOUT_FOLDER_LOCATION
        Dir1.Enabled = False
        Label7.Enabled = False
    Else
        Drive1.Enabled = True
        Dir1.Enabled = True
        Label7.Enabled = True
    End If

    mbFileNamePathChanged = True

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "chkSaveInAppPath_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub chkSiteCode_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    mbDisplaySiteCodeChanged = True

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "chkSiteCode_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub chkSplitGrid_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    mbSplitGridChanged = True

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "chkSplitGrid_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub chkStudyName_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    mbDisplayStudyNameChanged = True

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "chkStudyName_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'Mo 2/4/2007 MRC15022007
'--------------------------------------------------------------------
Private Sub chkUseStudyName_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    If chkUseStudyName.Value = 1 Then
        txtUserSpecName.Text = ""
        txtUserSpecName.Enabled = False
        Label8.Enabled = False
    Else
        txtUserSpecName.Enabled = True
        Label8.Enabled = True
    End If

    mbFileNameTextChanged = True

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "chkUseStudyName_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub chkVisitCycle_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    mbDisplayVisitCycleChanged = True

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "chkVisitCycle_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub cmdCancel_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    Unload Me

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdCancel_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub cmdOK_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    ActivateChanges
    
    Unload Me

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdOK_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'Mo 2/4/2007 MRC15022007
'--------------------------------------------------------------------
Private Sub Dir1_Change()
'--------------------------------------------------------------------

    On Error GoTo Errhandler

    mbFileNamePathChanged = True

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "Dir1_Change")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'Mo 2/4/2007 MRC15022007
'--------------------------------------------------------------------
Private Sub Drive1_Change()
'--------------------------------------------------------------------

    On Error GoTo DriveHandler
    
    
    Dir1.Path = Drive1.Drive
    
    mbFileNamePathChanged = True

Exit Sub
DriveHandler:
    Drive1.Drive = Dir1.Path
End Sub

'--------------------------------------------------------------------
Private Sub Form_Load()
'--------------------------------------------------------------------
'Mo 2/4/2007 MRC15022007
Dim sDrive As String

    On Error GoTo Errhandler
    
    Me.Icon = frmMenu.Icon
    FormCentre Me
 
    If gbDisplayStudyName = True Then
        chkStudyName.Value = 1
    Else
        chkStudyName.Value = 0
    End If
    
    If gbDisplaySiteCode = True Then
        chkSiteCode.Value = 1
    Else
        chkSiteCode.Value = 0
    End If

    If gbDisplayLabel = True Then
        chkLabel.Value = 1
    Else
        chkLabel.Value = 0
    End If
    
    If gbDisplayPersonId = True Then
        chkPersonId.Value = 1
    Else
        chkPersonId.Value = 0
    End If
    
    If gbDisplayVisitCycle = True Then
        chkVisitCycle.Value = 1
    Else
        chkVisitCycle.Value = 0
    End If
    
    If gbDisplayFormCycle = True Then
        chkFormCycle.Value = 1
    Else
        chkFormCycle.Value = 0
    End If
    
    If gbDisplayRepeatNumber = True Then
        chkRepeatNumber.Value = 1
    Else
        chkRepeatNumber.Value = 0
    End If
    
    If gbSplitGrid = True Then
        chkSplitGrid.Value = 1
    Else
        chkSplitGrid.Value = 0
    End If

    txtMissing.Text = gsSVMissing
    txtUnobtainable.Text = gsSVUnobtainable
    txtNotApplicable.Text = gsSVNotApplicable
    'Mo 1/11/2006 Bug 2795
    Select Case gnOutPutType
    Case eOutPutType.CSV
        optCSV.Value = True
    Case eOutPutType.Access
        optAccess.Value = True
    Case eOutPutType.SPSS
        optSPSS.Value = True
    Case eOutPutType.SAS, eOutPutType.SASColons
        optSAS.Value = True
    Case eOutPutType.STATA
        optSTATA.Value = True
    Case eOutPutType.MACROBD
        optMacroBD.Value = True
    Case eOutPutType.STATAStandardDates
        optSTATAStandardDates.Value = True
    End Select
    
    If gbOutputCategoryCodes Then
        optCodes.Value = True
    Else
        optValues.Value = True
    End If
    
    'Mo 2/6/2006 Bug 2737
    If gbUseShortCodes Then
        optShortCodes.Value = True
        txtShortCodeLength.Enabled = True
    Else
        optLongCodes.Value = True
        txtShortCodeLength.Enabled = False
    End If
    txtShortCodeLength.Text = gnShortCodeLength
    
    'Mo 30/5/2006 Bug 2668
    If gbExcludeLabel = True Then
        chkExcludeLabel.Value = 1
    Else
        chkExcludeLabel.Value = 0
    End If
    
    'Mo 1/11/2006 Bug 2795
    If gbSASInformatColons = True Then
        chkSASInformatColons.Value = 1
    Else
        chkSASInformatColons.Value = 0
    End If
    
    'Mo 2/4/2007 MRC15022007
    If gsFileNamePath = "" Then
        chkSaveInAppPath.Value = 1
        sDrive = Mid(gsOUT_FOLDER_LOCATION, 1, InStr(gsOUT_FOLDER_LOCATION, "\"))
        Drive1.Drive = sDrive
        Drive1.Enabled = False
        Dir1.Path = gsOUT_FOLDER_LOCATION
        Dir1.Enabled = False
    Else
        'Check that the specified folder exists, note that FolderExistence will create any folders that do not exist
        If FolderExistence(gsFileNamePath & "\") Then
            chkSaveInAppPath.Value = 0
            sDrive = Mid(gsFileNamePath, 1, InStr(gsFileNamePath, "\"))
            Drive1.Drive = sDrive
            Drive1.Enabled = True
            Dir1.Path = gsFileNamePath
            Dir1.Enabled = True
        Else
            'This code should never be reached because FolderExistence will create any folders that do not exist
            Call DialogInformation("Specified Output File Folder " & gsFileNamePath & " does not exist." & vbNewLine & "Folder being set to default of " & gsOUT_FOLDER_LOCATION)
            chkSaveInAppPath.Value = 1
            sDrive = Mid(gsOUT_FOLDER_LOCATION, 1, InStr(gsOUT_FOLDER_LOCATION, "\"))
            Drive1.Drive = sDrive
            Drive1.Enabled = False
            Dir1.Path = gsOUT_FOLDER_LOCATION
            Dir1.Enabled = False
            'reset the gsFileNamePath setting
            gsFileNamePath = ""
            'flag query as changed
            gbQueryChanged = True
        End If
    End If
    
    If gsFileNameText = "" Then
        chkUseStudyName.Value = 1
        txtUserSpecName.Enabled = False
    Else
        chkUseStudyName.Value = 0
        txtUserSpecName.Text = gsFileNameText
        txtUserSpecName.Enabled = True
    End If
    
    Select Case gsFileNameStamp
    Case "DATE"
        optDateStamp.Value = True
    Case "DATETIME"
        optDateTimeStamp.Value = True
    Case ""
        optNoStamp.Value = True
    End Select

    mbSVMissingChanged = False
    mbSVUnobtainableChanged = False
    mbSVNotApplicableChanged = False
    mbOutPutTypeChanged = False
    mbOutputCodeValuesChanged = False
    mbDisplayStudyNameChanged = False
    mbDisplaySiteCodeChanged = False
    mbDisplayLabelChanged = False
    mbDisplayPersonIdChanged = False
    mbDisplayVisitCycleChanged = False
    mbDisplayFormCycleChanged = False
    mbDisplayRepeatNumberChanged = False
    mbSplitGridChanged = False
    mbLongShortQuestionsChanged = False
    'Mo 30/5/2006 Bug 2668
    mbExcludeLabelChanged = False
    'Mo 1/11/2006 Bug 2795
    mbSASInformatColonsChanged = False
    'Mo 2/4/2007 MRC15022007
    mbFileNamePathChanged = False
    mbFileNameTextChanged = False
    mbFileNameStampChanged = False
    
    Call AssessOptions
    
Exit Sub
Errhandler:
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

'--------------------------------------------------------------------
Private Sub optAccess_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    Call AssessOptions
    
    mbOutPutTypeChanged = True

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "optAccess_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub optCodes_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    mbOutputCodeValuesChanged = True
    
Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "optCodes_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub optCSV_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    Call AssessOptions
    
    mbOutPutTypeChanged = True

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "optCSV_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'Mo 2/4/2007 MRC15022007
'--------------------------------------------------------------------
Private Sub optDateStamp_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    mbFileNameStampChanged = True
    
Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "optDateStamp_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'Mo 2/4/2007 MRC15022007
'--------------------------------------------------------------------
Private Sub optDateTimeStamp_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    mbFileNameStampChanged = True
    
Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "optDateTimeStamp_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub optLongCodes_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    mbLongShortQuestionsChanged = True
    
    'Mo 2/6/2006 Bug 2737
    txtShortCodeLength.Enabled = False
    txtShortCodeLength.Text = 8
    Label5.Enabled = False
    
Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "optLongCodes_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub optMacroBD_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    Call AssessOptions
    
    mbOutPutTypeChanged = True

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "optMacroBD_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'Mo 2/4/2007 MRC15022007
'--------------------------------------------------------------------
Private Sub optNoStamp_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    mbFileNameStampChanged = True
    
Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "optNoStamp_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub optSAS_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    Call AssessOptions
    
    mbOutPutTypeChanged = True

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "optSAS_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub optShortCodes_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    mbLongShortQuestionsChanged = True
    
    'Mo 2/6/2006 Bug 2737
    txtShortCodeLength.Enabled = True
    Label5.Enabled = True

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "optShortCodes_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub optSPSS_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    Call AssessOptions
    
    mbOutPutTypeChanged = True

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "optSPSS_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub optSTATA_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    Call AssessOptions
    
    mbOutPutTypeChanged = True

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "optSTATA_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub optSTATAStandardDates_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    Call AssessOptions
    
    mbOutPutTypeChanged = True

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "optSTATAStandardDates_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub optValues_Click()
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    mbOutputCodeValuesChanged = True
    
Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "optValues_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub txtMissing_Change()
'--------------------------------------------------------------------
Dim sText As String

    On Error GoTo Errhandler
    
    mbSVMissingChanged = True
    sText = txtMissing.Text
    If sText <> "" And sText <> "-" And sText <> "-1" And sText <> "-2" _
    And sText <> "-3" And sText <> "-4" And sText <> "-5" And sText <> "-6" _
    And sText <> "-7" And sText <> "-8" And sText <> "-9" Then
        txtMissing.Text = ""
        Call DialogInformation("Special values can only be negative numerics between -1 and -9", "Special Values")
    End If
    
    Call AssessOptions

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtMissing_Change")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub txtMissing_LostFocus()
'--------------------------------------------------------------------
Dim sText As String

    On Error GoTo Errhandler
    
    mbSVMissingChanged = True
    sText = txtMissing.Text
    If sText <> "" And sText <> "-1" And sText <> "-2" _
    And sText <> "-3" And sText <> "-4" And sText <> "-5" And sText <> "-6" _
    And sText <> "-7" And sText <> "-8" And sText <> "-9" Then
        txtMissing.Text = ""
        Call DialogInformation("Special values can only be negative numerics between -1 and -9", "Special Values")
    End If
    
    Call AssessOptions

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtMissing_LostFocus")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub txtNotApplicable_Change()
'--------------------------------------------------------------------
Dim sText As String

    On Error GoTo Errhandler
    
    mbSVNotApplicableChanged = True
    sText = txtNotApplicable.Text
    If sText <> "" And sText <> "-" And sText <> "-1" And sText <> "-2" _
    And sText <> "-3" And sText <> "-4" And sText <> "-5" And sText <> "-6" _
    And sText <> "-7" And sText <> "-8" And sText <> "-9" Then
        txtNotApplicable.Text = ""
        Call DialogInformation("Special values can only be negative numerics between -1 and -9", "Special Values")
    End If
    
    Call AssessOptions

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtNotApplicable_Change")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub txtNotApplicable_LostFocus()
'--------------------------------------------------------------------
Dim sText As String

    On Error GoTo Errhandler
    
    mbSVNotApplicableChanged = True
    sText = txtNotApplicable.Text
    If sText <> "" And sText <> "-1" And sText <> "-2" _
    And sText <> "-3" And sText <> "-4" And sText <> "-5" And sText <> "-6" _
    And sText <> "-7" And sText <> "-8" And sText <> "-9" Then
        txtNotApplicable.Text = ""
        Call DialogInformation("Special values can only be negative numerics between -1 and -9", "Special Values")
    End If
    
    Call AssessOptions

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtNotApplicable_LostFocus")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'Mo 2/6/2006 Bug 2737
'--------------------------------------------------------------------
Private Sub txtShortCodeLength_Change()
'--------------------------------------------------------------------
Dim sText As String

    On Error GoTo Errhandler
    
    mbLongShortQuestionsChanged = True
    sText = txtShortCodeLength.Text
    If sText <> "1" And sText <> "8" And sText <> "9" And sText <> "10" And sText <> "11" _
    And sText <> "12" And sText <> "13" And sText <> "14" And sText <> "15" _
    And sText <> "16" And sText <> "17" And sText <> "18" Then
        Call DialogInformation("Short Code length should be a value between 8 and 18")
        txtShortCodeLength.Text = gnShortCodeLength
    End If

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtShortCodeLength_Change")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'Mo 2/6/2006 Bug 2737
'--------------------------------------------------------------------
Private Sub txtShortCodeLength_LostFocus()
'--------------------------------------------------------------------
Dim sText As String

    On Error GoTo Errhandler
    
    mbLongShortQuestionsChanged = True
    sText = txtShortCodeLength.Text
    If sText <> "8" And sText <> "9" And sText <> "10" And sText <> "11" _
    And sText <> "12" And sText <> "13" And sText <> "14" And sText <> "15" _
    And sText <> "16" And sText <> "17" And sText <> "18" Then
        Call DialogInformation("Short Code length should be a value between 8 and 18")
        txtShortCodeLength.Text = gnShortCodeLength
    End If

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtShortCodeLength_LostFocus")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub txtUnobtainable_Change()
'--------------------------------------------------------------------
Dim sText As String

    On Error GoTo Errhandler
    
    mbSVUnobtainableChanged = True
    sText = txtUnobtainable.Text
    If sText <> "" And sText <> "-" And sText <> "-1" And sText <> "-2" _
    And sText <> "-3" And sText <> "-4" And sText <> "-5" And sText <> "-6" _
    And sText <> "-7" And sText <> "-8" And sText <> "-9" Then
        txtUnobtainable.Text = ""
        Call DialogInformation("Special values can only be negative numerics between -1 and -9", "Special Values")
    End If
    
    Call AssessOptions

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtUnobtainable_Change")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub txtUnobtainable_LostFocus()
'--------------------------------------------------------------------
Dim sText As String

    On Error GoTo Errhandler
    
    mbSVUnobtainableChanged = True
    sText = txtUnobtainable.Text
    If sText <> "" And sText <> "-1" And sText <> "-2" _
    And sText <> "-3" And sText <> "-4" And sText <> "-5" And sText <> "-6" _
    And sText <> "-7" And sText <> "-8" And sText <> "-9" Then
        txtUnobtainable.Text = ""
        Call DialogInformation("Special values can only be negative numerics between -1 and -9", "Special Values")
    End If
    
    Call AssessOptions

Exit Sub
Errhandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtUnobtainable_LostFocus")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
End Sub

'--------------------------------------------------------------------
Private Sub ActivateChanges()
'--------------------------------------------------------------------

    On Error GoTo Errhandler

    If mbSVMissingChanged Then
        gsSVMissing = txtMissing.Text
        'flag query as changed
        gbQueryChanged = True
    End If
    
    If mbSVUnobtainableChanged Then
        gsSVUnobtainable = txtUnobtainable.Text
        gbQueryChanged = True
    End If
    
    If mbSVNotApplicableChanged Then
        gsSVNotApplicable = txtNotApplicable.Text
        gbQueryChanged = True
    End If
    
    'Mo 1/11/2006 Bug 2795
    If mbSASInformatColonsChanged Then
        If chkSASInformatColons.Value = 1 Then
            gbSASInformatColons = True
        Else
            gbSASInformatColons = False
        End If
        gbQueryChanged = True
    End If
    
    If mbOutPutTypeChanged Then
        If optCSV.Value = True Then
            gnOutPutType = eOutPutType.CSV
        ElseIf optAccess.Value = True Then
            gnOutPutType = eOutPutType.Access
        ElseIf optSPSS.Value = True Then
            gnOutPutType = eOutPutType.SPSS
        ElseIf optSAS.Value = True Then
            If gbSASInformatColons = True Then
                gnOutPutType = eOutPutType.SASColons
            Else
                gnOutPutType = eOutPutType.SAS
            End If
        ElseIf optSTATA.Value = True Then
            gnOutPutType = eOutPutType.STATA
        ElseIf optMacroBD.Value = True Then
            gnOutPutType = eOutPutType.MACROBD
        ElseIf optSTATAStandardDates.Value = True Then
            gnOutPutType = eOutPutType.STATAStandardDates
        End If
        gbQueryChanged = True
    End If
    
    If mbDisplayStudyNameChanged Then
        If chkStudyName.Value = 1 Then
            gbDisplayStudyName = True
        Else
            gbDisplayStudyName = False
        End If
        gbQueryChanged = True
    End If
    
    If mbDisplaySiteCodeChanged Then
        If chkSiteCode.Value = 1 Then
            gbDisplaySiteCode = True
        Else
            gbDisplaySiteCode = False
        End If
        gbQueryChanged = True
    End If

    If mbDisplayLabelChanged Then
        If chkLabel.Value = 1 Then
            gbDisplayLabel = True
        Else
            gbDisplayLabel = False
        End If
        gbQueryChanged = True
    End If
    
    If mbDisplayPersonIdChanged Then
        If chkPersonId.Value = 1 Then
            gbDisplayPersonId = True
        Else
            gbDisplayPersonId = False
        End If
        gbQueryChanged = True
    End If
    
    If mbDisplayVisitCycleChanged Then
        If chkVisitCycle.Value = 1 Then
            gbDisplayVisitCycle = True
        Else
            gbDisplayVisitCycle = False
        End If
        gbQueryChanged = True
    End If
    
    If mbDisplayFormCycleChanged Then
        If chkFormCycle.Value = 1 Then
            gbDisplayFormCycle = True
        Else
            gbDisplayFormCycle = False
        End If
        gbQueryChanged = True
    End If
    
    If mbDisplayRepeatNumberChanged Then
        If chkRepeatNumber.Value = 1 Then
            gbDisplayRepeatNumber = True
        Else
            gbDisplayRepeatNumber = False
        End If
        gbQueryChanged = True
    End If
    
    If mbSplitGridChanged Then
        If chkSplitGrid.Value = 1 Then
            gbSplitGrid = True
        Else
            gbSplitGrid = False
        End If
        gbQueryChanged = True
    End If
    
    If mbOutputCodeValuesChanged Then
        If optCodes.Value = True Then
            gbOutputCategoryCodes = True
        Else
            gbOutputCategoryCodes = False
        End If
        'Changed Mo 11/7/2002
        'This setting will change the response for a category question when a query is run/re-run.
        'Given that gbOutputCategoryCodes and the contents of grdOutPut do not correspond until a
        'query is run/re-run the cmdSaveOutput button is disabled.
        frmMenu.cmdSaveOutPut.Enabled = False
        gbQueryChanged = True
    End If
    
    'Mo 2/6/2006 Bug 2737
    If mbLongShortQuestionsChanged Then
        If optShortCodes.Value = True Then
            gbUseShortCodes = True
            gnShortCodeLength = txtShortCodeLength.Text
        Else
            gbUseShortCodes = False
            gnShortCodeLength = 8
        End If
        'Having changed gbUseShortCodes the column headings in frmMenu.grdOutPut will not
        'correspond until a query is run/re-run, the cmdSaveOutput button is disabled.
        frmMenu.cmdSaveOutPut.Enabled = False
        gbQueryChanged = True
    End If
    
    'Mo 30/5/2006 Bug 2668
    If mbExcludeLabelChanged Then
        If chkExcludeLabel.Value = 1 Then
            gbExcludeLabel = True
        Else
            gbExcludeLabel = False
        End If
        gbQueryChanged = True
    End If
    
    'Mo 2/4/2007 MRC15022007
    If mbFileNamePathChanged Then
        If chkSaveInAppPath.Value = 1 Then
            gsFileNamePath = ""
        Else
            gsFileNamePath = Dir1.Path
        End If
        gbQueryChanged = True
    End If
    
    If mbFileNameTextChanged Then
        If chkUseStudyName.Value = 1 Then
            gsFileNameText = ""
        Else
            gsFileNameText = txtUserSpecName.Text
        End If
        gbQueryChanged = True
    End If
    
    If mbFileNameStampChanged Then
        If optDateStamp.Value = True Then
            gsFileNameStamp = "DATE"
        ElseIf optDateTimeStamp.Value = True Then
            gsFileNameStamp = "DATETIME"
        Else
            gsFileNameStamp = ""
        End If
        gbQueryChanged = True
    End If

Exit Sub
Errhandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "ActivateChanges")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'--------------------------------------------------------------------
Private Sub AssessOptions()
'--------------------------------------------------------------------
'The aim of this sub is to:-
'   Prevent "MACRO Batch Data Format" from being selected in any Special
'   Values have been setup.
'and vice versa
'   Prevent any Special Values being set if "MACRO Batch Data Format" has
'   been selected.
'--------------------------------------------------------------------
'Mo 1/11/2006 Bug 2795, this sub now enables/disables chkSASInformatColons
'   based on whether SAS output option has been selected
'--------------------------------------------------------------------

    On Error GoTo Errhandler
    
    If (txtMissing.Text = "") And (txtUnobtainable.Text = "") And (txtNotApplicable.Text = "") Then
        optMacroBD.Enabled = True
    Else
        optMacroBD.Value = False
        optMacroBD.Enabled = False
    End If
    
    If optMacroBD.Value = True Then
        txtMissing.Enabled = False
        txtUnobtainable.Enabled = False
        txtNotApplicable.Enabled = False
    Else
        txtMissing.Enabled = True
        txtUnobtainable.Enabled = True
        txtNotApplicable.Enabled = True
    End If
    
    'Mo 1/11/2006 Bug 2795, enable chkSASInformatColons if SAS ouput option is selected
    If optSAS.Value = True Then
        chkSASInformatColons.Enabled = True
    Else
        'clear hkSASInformatColons before disabling it
        chkSASInformatColons.Value = 0
        chkSASInformatColons.Enabled = False
    End If
        
Exit Sub
Errhandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "AssessOptions")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'Mo 2/4/2007 MRC15022007
'--------------------------------------------------------------------
Private Sub txtUserSpecName_Change()
'--------------------------------------------------------------------

    On Error GoTo Errhandler

    mbFileNameTextChanged = True
    
    If Not gblnValidString(txtUserSpecName.Text, valOnlySingleQuotes) Then
        Call DialogInformation("File name " & gsCANNOT_CONTAIN_INVALID_CHARS)
        txtUserSpecName.Text = ""
        Exit Sub
    ElseIf Not gblnValidString(txtUserSpecName.Text, valAlpha + valNumeric) Then
        Call DialogInformation("File name can only contain alphanumeric characters.")
        txtUserSpecName.Text = ""
        Exit Sub
    ElseIf Len(txtUserSpecName.Text) > 20 Then
        Call DialogInformation("File name can not be more than 20 characters")
        txtUserSpecName.Text = ""
        Exit Sub
    End If

Exit Sub
Errhandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "txtUserSpecName_Change")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub
