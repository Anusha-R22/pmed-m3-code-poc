VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MACRO Exchange"
   ClientHeight    =   5445
   ClientLeft      =   5505
   ClientTop       =   3645
   ClientWidth     =   7155
   Icon            =   "frmMenuExchange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   7155
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   4620
      Width           =   1215
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   1560
      TabIndex        =   17
      Top             =   4620
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5880
      TabIndex        =   19
      Top             =   4620
      Width           =   1200
   End
   Begin MSComctlLib.StatusBar sbrMenu 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   5100
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "UserKey"
            Object.ToolTipText     =   "Name of current user"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "RoleKey"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "UserDatabase"
            Object.ToolTipText     =   "Current user database"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab ssTabExchange 
      Height          =   4455
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   7858
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Site Utilities"
      TabPicture(0)   =   "frmMenuExchange.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmTrialSite"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Study Definition Utilities"
      TabPicture(1)   =   "frmMenuExchange.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmSDD"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Laboratory Utilities"
      TabPicture(2)   =   "frmMenuExchange.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraLab"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Subject Data Utilities"
      TabPicture(3)   =   "frmMenuExchange.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "frmPRD"
      Tab(3).ControlCount=   1
      Begin VB.Frame frmTrialSite 
         Height          =   3975
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   6795
         Begin VB.CommandButton cmdJavaScript 
            Caption         =   "Create Java Script (temp)"
            Height          =   375
            Left            =   2160
            TabIndex        =   22
            Top             =   3240
            Width           =   2430
         End
         Begin VB.CommandButton cmdSiteLab 
            Caption         =   "Laboratory Site Administration"
            Height          =   375
            Left            =   2160
            TabIndex        =   4
            Top             =   2280
            Width           =   2430
         End
         Begin VB.CommandButton CmdSiteAdmin 
            Caption         =   "Site Administration"
            Height          =   375
            Left            =   2160
            TabIndex        =   2
            Top             =   1320
            Width           =   2430
         End
         Begin VB.CommandButton CmdTrialSiteAdmin 
            Caption         =   "Study Site Administration"
            Height          =   375
            Left            =   2160
            TabIndex        =   3
            Top             =   1800
            Width           =   2430
         End
      End
      Begin VB.Frame frmSDD 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   20
         Top             =   360
         Width           =   6795
         Begin VB.CommandButton CmdTrialStatus 
            Caption         =   "Study Status"
            Height          =   375
            Left            =   2160
            TabIndex        =   8
            Top             =   2760
            Width           =   2430
         End
         Begin VB.CommandButton cmdExportStudy 
            Caption         =   "Export Study Definition"
            Height          =   375
            Left            =   2160
            TabIndex        =   5
            Top             =   1320
            Width           =   2430
         End
         Begin VB.CommandButton cmdGenerateHTML 
            Caption         =   "Generate HTML Forms"
            Height          =   375
            Left            =   2160
            TabIndex        =   7
            Top             =   2280
            Width           =   2430
         End
         Begin VB.CommandButton cmdImportStudy 
            Caption         =   "Import Study Definition"
            Height          =   375
            Left            =   2160
            TabIndex        =   6
            Top             =   1800
            Width           =   2430
         End
      End
      Begin VB.Frame frmPRD 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   18
         Top             =   360
         Width           =   6795
         Begin VB.CommandButton cmdExportTAB 
            Caption         =   "TAB Delimited Export"
            Height          =   375
            Left            =   2160
            TabIndex        =   14
            Top             =   2280
            Width           =   2430
         End
         Begin VB.CommandButton cmdImportPatientData 
            Caption         =   "Import Subject Data"
            Height          =   375
            Left            =   2160
            TabIndex        =   13
            Top             =   1800
            Width           =   2430
         End
         Begin VB.CommandButton cmdExportPatientData 
            Caption         =   "Export Subject Data"
            Height          =   375
            Left            =   2160
            TabIndex        =   12
            Top             =   1320
            Width           =   2430
         End
      End
      Begin VB.Frame fraLab 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   16
         Top             =   360
         Width           =   6795
         Begin VB.CommandButton cmdDistributeLab 
            Caption         =   "Distribute Laboratory Definition"
            Height          =   375
            Left            =   2160
            TabIndex        =   11
            Top             =   2280
            Width           =   2430
         End
         Begin VB.CommandButton cmdImportLab 
            Caption         =   "Import Laboratory Definition"
            Height          =   375
            Left            =   2160
            TabIndex        =   10
            Top             =   1800
            Width           =   2430
         End
         Begin VB.CommandButton cmdExportLab 
            Caption         =   "Export Laboratory Definition"
            Height          =   375
            Left            =   2160
            TabIndex        =   9
            Top             =   1320
            Width           =   2430
         End
      End
   End
   Begin VB.Timer tmrSystemIdleTimeout 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   240
      Top             =   0
   End
   Begin MSComDlg.CommonDialog dlgMenu 
      Left            =   780
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       frmMenuExchange.frm
'   Author:     Andrew Newbigging, April 1998
'   Purpose:    Main menu used in Macro Exchange.
'--------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------
'   Revisions:
'   1    Andrewn            21/11/97
'   2    Andrewn            27/11/97
'   3    Andrewn            26/01/98
'   4    Mo Morris          24/02/98
'   5    Andrew Newbigging  24/02/98
'   6    Andrew Newbigging  2/04/98
'   7    Joanne Lau         30/04/98
'   8    Joanne Lau         6/05/98
'   9   Andrew Newbigging   11/11/98
'       Buttons rearranged onto tabs
'       Mo Morris           18/2/99
'       Changes to what the Export/Import patient data buttons call.
'       Mo Morris           12/4/99
'       Additional button and call to Tab Delimited export (frmTABExport) added
'   10  PN  10/09/99    Upgrade from DAO to ADO and updated code to conform
'                       to VB standards doc version 1.0
'   NCJ 16 Sept 1999
'       Added InitialiseMe routine
'   PN  21/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   PN  22/09/99    Moved call to ExitMacro from Form_QueryUnload() to where it should
'                   be in Form_Unload()
'   PN  23/09/99    Added tmrSystemShutdown timer control to manage proper shutdown
'                   of all forms in system, including modally displayed forms.
'                   Amended CmdTrialSiteAdmin_Click()
'   Mo Morris   24/9/99
'                   StartUpPSS added to InitialiseMe and ShutDownPSS added to Form_Unload
'   Mo Morris   15/12/99    Property Get Mode (=gsEXCHANGE_MODE) added
'   Mo Morris   17/12/99    cmdAbout and cmdHelp added
'   Mo Morris   17/12/99    cmdAutoExport removed
'   NCJ 21 Dec 99 - Changed "trial" to "study" in captions etc.
'   NCJ 13/1/00 Update gsMACROUserGuidePath in InitialiseMe
'   Mo Morris   14/1/00     The following controls have been removed:-
'                           cmdImportOCViews, cmdExportOCBatch, frmOracleClinical
'                           together with tab number 4 on labMenu
'   NCJ 15 Jan 00   SR2125 Disable functions according to user's access rights
'   Mo Morris   30/3/00     Changes around making all the command buttons and frames
'           have a consistent size and position.
'           cmdImportAll removed
'   TA 05/08/2000   subclassing removed
'   WillC SR3492 24/5/00 Made Show modal to stop frmMenu showing when calling msgbox in frmtrialstatus in cmdTrialStatus
'   WillC SR3563   2/8/00   Changed the wording on the frame, tab and 2 buttons to say Subject instead of patient. No code changes
'   required.
'   TA 06/10/2000: Added LabSite Button and changed tab to ssTab
'   TA 06/10/2000: Move Import Laboratory code to here
'   DPH 18/04/2002 - Include ZIP files
'   ASH 10/11/2002 - Added call to InitialiseSettingsFile in modsettings.bas
'   ZA 24/09/2002 - Removed PSS calls as it is no longer in use
'   ZA 26/09/2002 - Added create java script button for temp purposes
'--------------------------------------------------------------------------------

Option Explicit
Option Compare Binary
Option Base 0

Private mbSystemLocked As Boolean

'---------------------------------------------------------------------
Public Sub InitialiseMe()
'---------------------------------------------------------------------
' Perform initialisations specific to Exchange
' Called from Main in MainMacroModule at startup
' NCJ 16 Sept 1999
'---------------------------------------------------------------------
  
    gsMACROUserGuidePath = gsMACROUserGuidePath & "EX\Contents.htm"
  
    ' NCJ 15 Jan 00 - Enable buttons according to user's rights
    If goUser.CheckPermission(gsFnImportPatData) Then
        cmdImportPatientData.Enabled = True
    Else
        cmdImportPatientData.Enabled = False
    End If
    If goUser.CheckPermission(gsFnExportPatData) Then
        cmdExportPatientData.Enabled = True
        cmdExportTAB.Enabled = True
    Else
        cmdExportPatientData.Enabled = False
        cmdExportTAB.Enabled = False
    End If
    If goUser.CheckPermission(gsFnImportStudyDef) Then
        cmdImportStudy.Enabled = True
    Else
        cmdImportStudy.Enabled = False
    End If
    If goUser.CheckPermission(gsFnDistribNewVersionOfStudyDef) Then
        cmdExportStudy.Enabled = True
    Else
        cmdExportStudy.Enabled = False
    End If
    
End Sub

''---------------------------------------------------------------------
'Private Sub cmdAutoExport_Click()
''---------------------------------------------------------------------
'Dim oExchange As clsExchange
'  On Error GoTo ErrHandler
'
'    Screen.MousePointer = vbHourglass
'
'    Set oExchange = New clsExchange
'    oExchange.AutoExportPRD "Auto"
'
'    Screen.MousePointer = vbDefault
'
'Exit Sub
'ErrHandler:
'  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
'                                                "cmdAutoExport_Click")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Unload frmMenu
'   End Select
'
'End Sub

'---------------------------------------------------------------------
Private Sub cmdDistributeLab_Click()
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
  
    Call frmExportLab.Display(True)
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "cmdDistributeLab_Click")
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
Private Sub cmdExit_Click()
'---------------------------------------------------------------------
' close the app
' the cleaning up will be handled in the query_unload event
'---------------------------------------------------------------------
  
    Unload Me
        
End Sub

'---------------------------------------------------------------------
Private Sub cmdExportLab_Click()
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
  
    Call frmExportLab.Display(False)
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "cmdExportLab_Click")
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
Private Sub cmdExportTAB_Click()
'---------------------------------------------------------------------
  
    frmTABExport.Show vbModal
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdExportPatientData_Click()
'---------------------------------------------------------------------
  On Error GoTo ErrHandler
  
    frmExportPatientData.Show vbModal
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "cmdExportPatientData_Click")
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
Private Sub cmdExportStudy_Click()
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
  
    frmExportStudyDefinition.Show vbModal
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "cmdExportStudy_Click")
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
Private Sub cmdGenerateHTML_Click()
'---------------------------------------------------------------------
    
    frmGenerateHTML.Show vbModal
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdHelp_Click()
'---------------------------------------------------------------------
' show the helpfile
' NCJ 13/1/00 - Real Help file hook added (activate when files exist)
' NCJ 20/1/00 - Activated Help file call
'---------------------------------------------------------------------

    'Call ShowDocument(Me.hWnd, gsMACROUserGuidePath)
            
    'REM 07/12/01 - New Call to MACRO Help
    Call MACROHelp(Me.hWnd, App.Title)
             
End Sub

'---------------------------------------------------------------------
Private Sub cmdImportAll_Click()
'---------------------------------------------------------------------

'Dim oMessages As clsMessages
    '
    'Screen.MousePointer = vbHourglass
    '
    'Set oMessages = New clsMessages
    'oMessages.ImportIncomingMessages (gsIN_FOLDER_LOCATION)
    '
    'Screen.MousePointer = vbDefault

End Sub

'---------------------------------------------------------------------
Private Sub cmdImportLab_Click()
'---------------------------------------------------------------------
' DPH 18/04/2002 - Include ZIP files
'---------------------------------------------------------------------
Dim sNextLDDFile As String
Dim oExchange As clsExchange
Dim sImportFile As String

    On Error GoTo ErrHandler
  
    sImportFile = gsIN_FOLDER_LOCATION & "*.*"
    If CMDialogOpen(dlgMenu, "Select Laboratory Definition Import File", sImportFile, "Laboratory Definition Import Files (*.cab;*.zip)|*.cab;*.zip") Then
  
        If DialogQuestion("Are you sure you wish to import laboratory definition import file" & vbCrLf & sImportFile) = vbYes Then
            HourglassOn
            
            Set oExchange = New clsExchange
        
            'Unpack the CAB file into an .ldd file into directory AppPath/CabExtract
            oExchange.ImportLDDCAB (sImportFile)
            
            'loop through the extracted files and import them into Macro
            sNextLDDFile = Dir(gsCAB_EXTRACT_LOCATION & "*.ldd")
            If sNextLDDFile = "" Then
                'no ldd files in cab
                Call MsgBox("No laboratory definition file found" + vbNewLine + "Import aborted.", , "Import Laboratory")
            Else
                'display status form
                Call frmStatus.Start(GetApplicationTitle, "Importing laboratory definition " & StripFileNameFromPath(sNextLDDFile) & "...", False)
                
                Select Case oExchange.ImportLDD(gsCAB_EXTRACT_LOCATION & sNextLDDFile)
                Case ExchangeError.EmptyFile
                    Call frmStatus.Finish
                    Call MsgBox(sImportFile & " is empty." + vbNewLine + "Import aborted.", , "Import Laboratory")
                Case ExchangeError.Invalid
                    Call frmStatus.Finish
                    Call MsgBox(sImportFile & " is not a valid laboratory definition file." + vbNewLine + "Import aborted.", , "Import Laboratory")
                Case ExchangeError.DirectoryNotFound
                    Call frmStatus.Finish
                    Call MsgBox(sImportFile & " does not exist." + vbNewLine + "Import aborted.", , "Import Laboratory")
                Case ExchangeError.Success
                    Call frmStatus.Finish
                    Call MsgBox(sNextLDDFile & " imported.", , "Import Laboratory")
                Case Else
                    Call frmStatus.Finish
                    Call MsgBox("Unexpected error. Import aborted", , "Import Laboratory")
                End Select
            End If
            
            HourglassOff
        End If
    End If
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "cmdImportLab_Click")
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
Private Sub cmdImportPatientData_Click()
'---------------------------------------------------------------------
  On Error GoTo ErrHandler
  
    frmImportPatientData.Show vbModal
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "cmdImportPatientData_Click")
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
Private Sub cmdImportStudy_Click()
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
  
    frmImportStudyDefinition.ImportType = SDDImportType.MACRO
    frmImportStudyDefinition.Show vbModal
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "cmdImportStudy_Click")
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
Private Sub cmdJavaScript_Click()
'---------------------------------------------------------------------
'Temporary button placement, calls all functions from modJScript
'---------------------------------------------------------------------
    
    'CreateTrialList
    'CreateSitesList
    Call CreateVisitList(goUser)
    Call CreateEFormsList(goUser)
    Call CreateQuestionsList(goUser)
    Call CreateUsersList(goUser)
    
End Sub

'---------------------------------------------------------------------
Private Sub CmdSiteAdmin_Click()
'---------------------------------------------------------------------
'---------------------------------------------------------------------
       
    ' NCJ 31/5/00 - Show as modal (same as others)
    FrmSiteAdmin.Show vbModal
        
End Sub

'---------------------------------------------------------------------
Private Sub cmdSiteLab_Click()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
  
    Call frmTrialSiteAdmin.Display(eDisplayType.DisplaySitesByLab)
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdSiteLab_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select

End Sub

'---------------------------------------------------------------------
Private Sub CmdTrialSiteAdmin_Click()
'---------------------------------------------------------------------
' REVISIONS
' DPH 13/08/2002 - Trial Site Administration form change for Study Versioning
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    ' PN 23/09/99
    ' show form modally to prevent attempt to load same form twice
    ' DPH 13/08/2002 - Show Study Versioning form
    'Call frmTrialSiteAdmin.Display(DisplaySitesByTrial)
    Call frmTrialSiteAdminVersioning.Display(DisplaySitesByTrial)
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "CmdTrialSiteAdmin_Click")
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
Private Sub CmdTrialStatus_Click()
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    'WillC SR3492 Made Show modal to stop frmMenu showing when calling msgbox in frmtrialstatus
    frmTrialStatus.Show vbModal
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "CmdTrialStatus_Click")
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
Private Sub cmdAbout_Click()
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    frmAbout.Show
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                                "cmdAbout_Click")
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
Private Sub cmdUserSite_Click()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
  
    frmTrialSiteAdmin.Display (eDisplayType.DisplaySitesByUser)
    
Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "cmdUserSite_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select


End Sub

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------
    
    FormCentre Me

End Sub

'---------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'---------------------------------------------------------------------

    Call ExitMACRO
   
End Sub


'---------------------------------------------------------------------
Private Sub tmrSystemIdleTimeout_Timer()
'---------------------------------------------------------------------
' nb This timer event should not occur unless DevMode = 0
' when the timer goes off it must be time to lock the system
'  it prompts the user to enter the password or wxit MACRO
' the system is then either closed in a controlled way or resets the timer
' NCJ 17/3/00 - Tidied up and simplified (SR 3015)
'TA 27/04/2000: new timeout handling
'---------------------------------------------------------------------
    'new timeout handling
    glSystemIdleTimeoutCount = glSystemIdleTimeoutCount + 1
    If glSystemIdleTimeout = glSystemIdleTimeoutCount Then
        ' set the couter to 0 and disable the timer until the user logs in
        glSystemIdleTimeoutCount = 0
        tmrSystemIdleTimeout.Enabled = False
        If frmTimeOutSplash.Display Then
            'password correctly entered
            tmrSystemIdleTimeout.Enabled = True
        Else
            'exit MACRO chosen
            ' unload all forms and exit
            Call UnloadAllForms
        End If
    End If
    
End Sub

'---------------------------------------------------------------------
Public Sub CheckUserRights()
'---------------------------------------------------------------------

End Sub

'---------------------------------------------------------------------
Public Property Get Mode() As String
'---------------------------------------------------------------------

    Mode = gsEXCHANGE_MODE

End Property


'---------------------------------------------------------------------
Public Function ForgottenPassword(sSecurityCon As String, sUsername As String, sPassword As String, ByRef sErrMsg As String) As eDTForgottenPassword
'---------------------------------------------------------------------
'REM 06/12/02
'---------------------------------------------------------------------

    'dummy routine

End Function
