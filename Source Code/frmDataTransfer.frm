VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{0E56FD71-943D-11D2-BA66-0040053687FE}#1.0#0"; "DartWeb.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDataTransfer 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MACRO Data Transfer"
   ClientHeight    =   8355
   ClientLeft      =   6990
   ClientTop       =   2940
   ClientWidth     =   9030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H8000000B&
      Caption         =   "&Start Transfer"
      Default         =   -1  'True
      Height          =   345
      Left            =   3720
      TabIndex        =   10
      Top             =   7935
      Width           =   1125
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1800
      Top             =   7860
   End
   Begin MSComctlLib.ProgressBar prgProgress 
      Height          =   315
      Left            =   540
      TabIndex        =   3
      Top             =   1620
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H8000000B&
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   7815
      TabIndex        =   2
      Top             =   7950
      Width           =   1125
   End
   Begin MSComCtl2.Animation anmavi 
      Height          =   615
      Left            =   3240
      TabIndex        =   1
      Top             =   60
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1085
      _Version        =   393216
      BackColor       =   -2147483637
      FullWidth       =   93
      FullHeight      =   41
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   5385
      Left            =   60
      TabIndex        =   11
      Top             =   2460
      Width           =   8895
      ExtentX         =   15690
      ExtentY         =   9499
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin DartWebCtl.Http Http1 
      Left            =   300
      OleObjectBlob   =   "frmDataTransfer.frx":0000
      Top             =   7860
   End
   Begin VB.Image imgError 
      Height          =   375
      Left            =   3900
      Picture         =   "frmDataTransfer.frx":00A0
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblAVIExist 
      BackColor       =   &H8000000B&
      Height          =   195
      Left            =   3300
      TabIndex        =   9
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblPercentage 
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   7920
      TabIndex        =   8
      Top             =   1620
      Width           =   495
   End
   Begin VB.Image imgArwR 
      Height          =   480
      Left            =   6240
      Picture         =   "frmDataTransfer.frx":010F
      Top             =   555
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape shpCircle2 
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   1860
      Shape           =   3  'Circle
      Top             =   690
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape shpCircle1 
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   6480
      Shape           =   3  'Circle
      Top             =   690
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      Visible         =   0   'False
      X1              =   2280
      X2              =   6300
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   5280
      TabIndex        =   7
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblEstTime 
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   1080
      TabIndex        =   6
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label lblConnecStatus 
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   1080
      TabIndex        =   5
      Top             =   1260
      Width           =   6135
   End
   Begin VB.Label lblServer 
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   6480
      TabIndex        =   4
      Top             =   60
      Width           =   1095
   End
   Begin VB.Image imgMACRO2 
      Height          =   795
      Left            =   6765
      Picture         =   "frmDataTransfer.frx":04A0
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Image imgMACRO1 
      Height          =   795
      Left            =   1020
      Picture         =   "frmDataTransfer.frx":0B02
      Stretch         =   -1  'True
      Top             =   360
      Width           =   795
   End
   Begin VB.Label lblSite 
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   1020
      TabIndex        =   0
      Top             =   60
      Width           =   1200
   End
   Begin VB.Image imgArwL 
      Height          =   480
      Left            =   1800
      Picture         =   "frmDataTransfer.frx":1164
      Top             =   540
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmDataTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 1998-2000. All Rights Reserved
'   File:       frmDataTransfer.frm
'   Author:     Richard Meinesz: April 2002
'   Purpose:    Form to inform the user during Data Transfer
'--------------------------------------------------------------------------------
'REVISONS:
' DPH 16/04/2002 -  Integrated code from frmExchangeStatus
'                   Added new HTTP contol (Dart Web)
'                   Added New status messages
' DPH 01/05/2002 - Replaced remaining GetStringResponseHTTP with PostDataHTTP calls
'                  Improved error handling on handling response data (DownloadMessages , DownloadMIMessages)
' DPH 16/05/2002 - Added command line data transfer with conditional compilation arguements
' DPH 10/06/2002 - Added information to Sending Notes/SDVs/Discrepancies CBBL 2.2.14.4
' ZA 24/06/2002  - Fixed bug 10 in build 2.2.14
' DPH 11/07/2002 - CBBL 2.2.19.31 & 2.2.19.34 DownLoadMessages - Close open recordset causing timeout / transaction problems in SQL Server
' DPH 27/08/2002 - Study Versioning Changes in DownloadMessages
' ZA  24/09/2002 - Remove code that uses Protocol storage
' RS 01/10/2002 - Added Timezone support
' NCJ 2 Oct 02 - Use more accurate IMedNow instead of CDbl(Now)
' TA 07/10/2002: Upgraded vb5 progress bar to vb6 progress bar
' NCJ 16 Oct 02 - Minor changes due to changes to MIMsg enumerations
' NCJ 19 Dec 02 - Added LFMessage stuff
' NCJ 3 Jan 02 - Added QuestionId to LFMessages
' NCJ 14 Jan 03 - We no longer ask Server to process LFMessages here (it's done in AutoImport)
' NCJ 20 Jan 03 - Added locking for receipt of LF Messages
' NCJ 22 Jan 03 - Corrected Connection String parameter for updating MIMsg status
' DPH 27/01/2003 - Added Report Transfer functionality
' REM 07/04/03 - added web control to display datatransfer log
' REM 01/12/03 - In routine SendMIMessages place RemoveNull around MIMessageText field
' NCJ 10 June 04 - Bug 2296 - Must reset oLFMsg each time round the loop in DownLoadLFMessages
' DPH 21/01/2005 - Added Pdu functionality to Data Transfer (PDU2300)
' DPH 22/02/2005 - Bug 2534 - check for pdu filename and delete if exists on download
' TA  18/08/2005 - Read timeout from settings file
' ic 20/11/2006 added check for ie7 format url
'--------------------------------------------------------------------------------

Option Explicit

Public Enum eConnectionStatus
    csConnecting = 1
    csConnectionEstablished = 2
    csDownloadingFromServer = 3
    csSendingDataToServer = 4
    csTransferComplete = 5
    csTransferCancelled = 6
    csSendingSystemDataToServer = 7
    csDownloadingSystemDataFromServer = 8
    csDownloadingReportFiles = 9
    csError = 10
    csSendingPduDataToServer = 11
    csDownloadingPduDataFromServer = 12
End Enum

Public Enum eTransfer
    trConnecting = 0
    trStepStarted = 1
    trTransferData = 2
    trStepCompleted = 3
    trTransferFinished = 4
    trTransferCancelled = 5
    trTransferError = 6
End Enum

Private msMACROServerDesc As String
Private msHTTPAddress As String
Private msSite As String
Private msUser As String
Private msPassword As String

Private Const msMessageListURL = "Exchange_Message_List.asp"
Private Const msGetNextMessageURL = "Exchange_Get_Next_Message.asp"
Private Const msNoMessagesURL = "exchange_no_messages.htm"
Private Const msGetStudyDefinitionURL = "exchange_get_study_definition.asp"
Private Const msReceiveDataURL = "exchange_receive_data.asp"
Private Const msReceiveData1URL = "exchange_receive_data1.asp"
Private Const msReceiveData2URL = "exchange_receive_data2.asp"
Private Const msReceiveData3URL = "exchange_receive_data3.asp"
Private Const msMIMessageListURL = "Exchange_MIMessage_List.asp"
Private Const msGetNextMIMessageURL = "Exchange_Get_Next_MIMessage.asp"
Private Const msReceiveMIMessageURL = "Exchange_receive_MIMessage.asp"
Private Const msReceiveLaboratoryURL = "Exchange_receive_Laboratory.asp"
Private Const msRegistrationURL = "Registration.asp"
Private Const msDataIntegrityURL = "DataIntegrityCheck.asp"
Private Const msDownloadInfoURL = "exchange_download_info.asp"
'REM 03/04/03 - test connection
Private Const msTEST_CONNECTION_URL = "Test_Connection.asp"

' NCJ 19 Dec 02 - LFMessage ASPs
Private Const msGetNextLFMessageURL = "Exchange_Get_Next_LFMessage.asp"
Private Const msReceiveLFMessageURL = "Exchange_Receive_LFMessage.asp"
Private Const msLFMessageListURL = "Exchange_LFMessage_List.asp"
Private Const msProcessLFMessagesURL = "Process_LFMessages.asp"

'REM 25/11/02 - System Data transfer asp pages
Private Const msCHECK_SYSTEM_MESSAGES_URL = "Check_System_Messages.asp"
Private Const msGET_SYSTEM_MESSAGES_URL = "Get_System_Messages.asp"
Private Const msWRITE_SYSTEM_MESSAGES_URL = "Write_System_Messages.asp"
Private Const msSYSTEM_MESSAGE_LIST_URL = "System_Message_List.asp"
Private Const msFORGOTTEN_PASSWORD_URL = "Forgotten_Password.asp"

' DPH 24/01/2003 - System report transfer asp page
Private Const msGetModifiedReportFiles = "Get_Report_Files.asp"

' DPH 21/01/2005 - PDU transfer asp pages
Private Const msPduMessageListURL = "Pdu_Message_List.asp"
Private Const msGetNextPduMessageURL = "Pdu_Next_Message.asp"
Private Const msDownloadPduFileURL = "Pdu_File_Download.asp"

Private Const msMIMESSAGE_DATA_INTEGRITY_ERROR = "MIMessage data integrity information could not be sent to "
Private Const msSUBJECT_DATA_INTEGRITY_ERROR = "Subject data integrity information could not be sent to "

' Constants for returned statuses from ASP pages
Private Const msASP_RETURN_SUCCESS = "SUCCESS"
Private Const msASP_RETURN_CHECKSUM_ERROR = "Error CAB Failed Checksum"

'REM 07/04/03 -  Data Transfer Web control tab stops
Private Const msTAB = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
Private Const msHALFTAB = "&nbsp;&nbsp;"

Private Const msBULLET_POINT = "<img src='../img/xferbullet.gif'>"
Private Const msSTEP_SUCCESS = "<img src='../img/ico_ok.gif'>"
Private Const msSTEP_FAIL = "<img src='../img/ico_error_perm.gif'>"
Private Const msTRANSFER_ERROR = "<img src='../img/ico_error_int.gif'>"

Private oExchange As clsExchange

Private mbCancelled As Boolean

' Variables to hold information about data to transfer
Private mlNoTransSubjects As Long
Private mlNoTransLockFreeze As Long
Private mlNoTransMIMessageDown As Long
Private mlNoTransMIMessageUp As Long
Private mlNoTransStudyUpdate As Long
Private mlNoTransLaboratoryDown As Long
Private mlNoTransLaboratoryUp As Long
Private mlNoTransStudyStatus As Long
Private mlNoSysMessagesDownload As Long
Private mlNoSysMessagesUpload As Long

Private mlNoTransLFMessageDown As Long
Private mlNoTransLFMessageUp As Long

Private mlNoTransReportFiles As Long

Private mDataTransferTime As clsDataTransferTime

' DPH 01/05/2002 - Error
Private mbTransferOK As Boolean
Private mbConnectionOK As Boolean

'TA store timeout from settings file
Private mnDataTransferTimeout As Integer

'--------------------------------------------------------------------------------
Public Sub Display(sSite As String, Optional nMax As Integer = 0)
'--------------------------------------------------------------------------------
' REM 10/04/02
' Display call
'--------------------------------------------------------------------------------
' DPH 29/04/2002 - Initialise Cancel button as close
'--------------------------------------------------------------------------------
    
    mbCancelled = False
    
    FormCentre Me
    
    'set the graphical interface on display
    Line1.Visible = False
    shpCircle1.Visible = False
    shpCircle2.Visible = False
    imgArwR.Visible = False
    imgArwL.Visible = False
    imgError.Visible = False
    imgMACRO2.Visible = False
    ' DPH 16/04/2002 - Enable animation control
    anmavi.Visible = True
    
    lblPercentage.Caption = ""
    lblEstTime.Caption = ""
    lblTime.Caption = ""
    
    lblSite.Caption = "Site: " & sSite
    
    'If Max for progress bar is set hen initalise the progress bar
    If nMax <> 0 Then
        Call InitProgressBar(nMax)
    End If

    'Me.Show vbModal
    Me.cmdCancel.Caption = "&Close"

    mnDataTransferTimeout = CInt(GetMACROSetting(MACRO_SETTING_DATATRANSFER_TIMEOUT, "60"))
End Sub

'--------------------------------------------------------------------------------
Private Sub SetConnectionStatus(nConnect As eConnectionStatus, Optional sMessage As String = "")
'--------------------------------------------------------------------------------
' REM 10/04/02
' Routine to set the graphical status depending on the connection status
'--------------------------------------------------------------------------------
' REVISIONS
' 16/04/2002 - Refresh form added
'REM 03/04/03 - added start/stop avi parameter
'--------------------------------------------------------------------------------
Dim sFileName As String
Dim sErr As String
Dim lErrNo As Long

    On Error GoTo StatusError
    
    'different connection status
    Select Case nConnect
    
    Case csConnecting
    
        'Start the avi
        On Error Resume Next
        sFileName = App.Path & "\MACRODataTransfer.avi"
        anmavi.Open sFileName
        anmavi.Play
        sErr = Err.Description
        lErrNo = Err.Number
        Err.Clear
        'check if avi file is present
        If lErrNo <> 0 Then
            lblAVIExist.Caption = "Transfer avi not found!"
        End If
        
        On Error GoTo StatusError
        
        If sMessage = "" Then
            lblConnecStatus.Caption = "Connecting to server"
        Else
            lblConnecStatus.Caption = sMessage
        End If
  
    Case csConnectionEstablished
        Line1.Visible = True
        shpCircle1.Visible = True
        shpCircle2.Visible = True
        imgMACRO2.Visible = True
        lblServer.Caption = "MACRO Server"
        
        If sMessage = "" Then
            lblConnecStatus.Caption = "Connection to server established"
        Else
            lblConnecStatus.Caption = sMessage
        End If
        
    Case csDownloadingFromServer
        Line1.Visible = True
        shpCircle1.Visible = False
        shpCircle2.Visible = False
        imgArwL.Visible = True
        imgMACRO2.Visible = True
        lblServer.Caption = "MACRO Server"
        
        If sMessage = "" Then
            lblConnecStatus.Caption = "Downloading messages from server"
        Else
            lblConnecStatus.Caption = sMessage
        End If
    
    Case csSendingDataToServer
        Line1.Visible = True
        shpCircle1.Visible = False
        shpCircle2.Visible = False
        imgArwL.Visible = False
        imgArwR.Visible = True
        imgMACRO2.Visible = True
        lblServer.Caption = "MACRO Server"
        
        If sMessage = "" Then
            lblConnecStatus.Caption = "Sending data to server"
        Else
            lblConnecStatus.Caption = sMessage
        End If
    
    Case csDownloadingReportFiles
        ' DPH 27/01/2003 Added Report File Transfer
        Line1.Visible = True
        shpCircle1.Visible = False
        shpCircle2.Visible = False
        imgArwL.Visible = False
        imgArwR.Visible = True
        imgMACRO2.Visible = True
        lblServer.Caption = "MACRO Server"
        
        If sMessage = "" Then
            lblConnecStatus.Caption = "Downloading report files from server"
        Else
            lblConnecStatus.Caption = sMessage
        End If
        
    Case csSendingSystemDataToServer
        Line1.Visible = True
        shpCircle1.Visible = False
        shpCircle2.Visible = False
        imgArwL.Visible = False
        imgArwR.Visible = True
        imgMACRO2.Visible = True
        lblServer.Caption = "MACRO Server"
        
        If sMessage = "" Then
            lblConnecStatus.Caption = "Sending system messages to server"
        Else
            lblConnecStatus.Caption = sMessage
        End If
    
    Case csDownloadingSystemDataFromServer
        Line1.Visible = True
        shpCircle1.Visible = False
        shpCircle2.Visible = False
        imgArwL.Visible = True
        imgMACRO2.Visible = True
        lblServer.Caption = "MACRO Server"
        
        If sMessage = "" Then
            lblConnecStatus.Caption = "Downloading system messages from server"
        Else
            lblConnecStatus.Caption = sMessage
        End If
        
    Case csDownloadingPduDataFromServer
        Line1.Visible = True
        shpCircle1.Visible = False
        shpCircle2.Visible = False
        imgArwL.Visible = True
        imgMACRO2.Visible = True
        lblServer.Caption = "MACRO Server"
        
        If sMessage = "" Then
            lblConnecStatus.Caption = "Downloading PDU messages from server"
        Else
            lblConnecStatus.Caption = sMessage
        End If
    
    Case csSendingPduDataToServer
        Line1.Visible = True
        shpCircle1.Visible = False
        shpCircle2.Visible = False
        imgArwL.Visible = True
        imgMACRO2.Visible = True
        lblServer.Caption = "MACRO Server"
        
        If sMessage = "" Then
            lblConnecStatus.Caption = "Launching PDU upload process"
        Else
            lblConnecStatus.Caption = sMessage
        End If
    
    Case csTransferComplete
        Line1.Visible = False
        shpCircle1.Visible = False
        shpCircle2.Visible = False
        imgArwL.Visible = False
        imgArwR.Visible = False
        imgMACRO2.Visible = True
        lblServer.Caption = "MACRO Server"
        
        ' DPH 16/04/2002 - stop avi
        anmavi.Stop
        anmavi.Visible = False
        
        If sMessage = "" Then
            lblConnecStatus.Caption = "Transfer complete"
        Else
            lblConnecStatus.Caption = sMessage
        End If
        
        cmdCancel.Caption = "&Close"
        
        lblEstTime.Caption = ""

    ' DPH 29/04/2002 - Transfer Cancelled
    Case csTransferCancelled
        Line1.Visible = False
        shpCircle1.Visible = False
        shpCircle2.Visible = False
        imgArwL.Visible = False
        imgArwR.Visible = False
        imgMACRO2.Visible = True
        lblServer.Caption = "MACRO Server"
        
        anmavi.Stop
        anmavi.Visible = False
        
        If sMessage = "" Then
            lblConnecStatus.Caption = "Transfer cancelled"
        Else
            lblConnecStatus.Caption = sMessage
        End If
        
        cmdCancel.Caption = "&Close"
        
        lblEstTime.Caption = ""
        lblTime.Caption = ""

    Case csError
        Line1.Visible = True
        shpCircle1.Visible = True
        shpCircle2.Visible = True
        imgArwL.Visible = False
        imgArwR.Visible = False
        imgMACRO2.Visible = False
        
        anmavi.Stop
        imgError.Visible = True
        lblServer.Caption = ""
        
        If sMessage = "" Then
            lblConnecStatus.Caption = "Transfer error!"
        Else
            lblConnecStatus.Caption = sMessage
        End If
        
        mbTransferOK = False
        
    End Select

    Me.Refresh
    
Exit Sub
StatusError:

End Sub

'--------------------------------------------------------------------------------
Public Sub InitProgressBar(nMax As Integer, Optional nMin As Integer = 1)
'--------------------------------------------------------------------------------
' REM 11/04/02
' Init progress bar
'--------------------------------------------------------------------------------
    
    With prgProgress
        .Value = nMin
        .Max = nMax
        .Min = nMin
    End With
    
    lblEstTime.Caption = ""
    lblTime.Caption = ""
    lblPercentage.Caption = ""
    
End Sub

'--------------------------------------------------------------------------------
Public Sub SetProgressBar(nValue As Integer, sTime As String)
'--------------------------------------------------------------------------------
' REM 10/04/02
' Set the progress bar value
'--------------------------------------------------------------------------------
' DPH 19/04/2002 - Include Data Transfer Time class
'--------------------------------------------------------------------------------
Dim lSecsRemaining As Long
Dim lMinsRemaining As Long

    With prgProgress
        .Value = nValue
        lblPercentage.Caption = Int(100 * (nValue / .Max)) & "%"
        If sTime <> "" Then
            lblEstTime.Caption = "Estimated time remaining"
            lblTime.Caption = sTime
        Else
            If .Value = .Max Then
                lblEstTime.Caption = ""
                lblTime.Caption = ""
            Else
                ' use DataTransferTime class
                lblEstTime.Caption = "Estimated time remaining"
                lSecsRemaining = mDataTransferTime.RemainingTime
                If lSecsRemaining > 60 Then
                    lMinsRemaining = lSecsRemaining \ 60
                    lblTime.Caption = lMinsRemaining & " minutes"
                Else
                    lblTime.Caption = lSecsRemaining & " seconds"
                End If
            End If
        End If
        Me.Refresh
    End With
    
End Sub

'--------------------------------------------------------------------------------
Public Sub SetTextBoxDisplay(sText As String, nTransfer As eTransfer, Optional bStatus As Boolean = True)
'--------------------------------------------------------------------------------
' REM 10/04/02
' Routine that writes messages to the text box.
'--------------------------------------------------------------------------------
' REVISIONS
' DPH 01/05/2002 - Optional parameter to give success / fail message
'--------------------------------------------------------------------------------
    
    Select Case nTransfer
    
    'Attemptingto connect to server
    Case trConnecting
        WebBrowser.Document.Write "<b>" & msHALFTAB & msBULLET_POINT & msHALFTAB & lblConnecStatus.Caption & "</b>" & "<br>" & msTAB & msTAB & "Started at " & Now & _
                "<br>" & msTAB & msTAB & sText & "<br>"
    
    'Started a transfer step
    Case trStepStarted
        WebBrowser.Document.Write "<b>" & msHALFTAB & msBULLET_POINT & msHALFTAB & lblConnecStatus.Caption & "</b>" & "<br>" & msTAB & msTAB & "Started at " & Now & "<br>" & msTAB & msTAB & Replace(sText, Chr(13) & Chr(10), "<br>" & msTAB & msTAB) & "<br>"
        
    'Transfering data
    Case trTransferData
        WebBrowser.Document.Write msTAB & msTAB & Replace(sText, Chr(13) & Chr(10), "<br>" & msTAB & msTAB) & "<br>"
    
    'Step completed
    Case trStepCompleted
        If bStatus Then
            WebBrowser.Document.Write msTAB & msTAB & sText & "<br>" & msTAB & msHALFTAB & msSTEP_SUCCESS & "Completed successfully at " & Now & "<p>"
        Else
            WebBrowser.Document.Write msTAB & msTAB & sText & "<br>" & msTAB & msHALFTAB & msSTEP_FAIL & "Completed with errors at " & Now & "<p>"
        End If
    
    'Data Transfer completed
    Case trTransferFinished
        If bStatus Then
            WebBrowser.Document.Write "<b>" & msHALFTAB & msBULLET_POINT & msHALFTAB & lblConnecStatus.Caption & "</b>" & "<br>" & msTAB & msHALFTAB & msSTEP_SUCCESS & "Completed successfully at " & Now & "<p>"
        Else
            WebBrowser.Document.Write "<b>" & msHALFTAB & msBULLET_POINT & msHALFTAB & lblConnecStatus.Caption & "</b>" & "<br>" & msTAB & msHALFTAB & msSTEP_FAIL & "Completed with errors at " & Now & "<p>"
        End If

        
    'Data Transfer Cancelled
    Case trTransferCancelled
        WebBrowser.Document.Write "<b>" & msHALFTAB & msSTEP_FAIL & msHALFTAB & lblConnecStatus.Caption & "</b>" & "<br>" & msTAB & msTAB & "Data Transfer Cancelled at " & Now & "<br>" & msTAB & msTAB & sText & "<p>"
    
    'Error!
    Case trTransferError
        WebBrowser.Document.Write "<p>" & "<b>" & msTRANSFER_ERROR & msHALFTAB & "<font color='red'>" & "Data Transfer Error" & "</font>" & "</b>" & "<br>" & msTAB & msTAB & sText
        
    End Select
    
    'move cursor to end of text box
    WebBrowser.Document.parentWindow.Scroll 0, 1000000000
    
End Sub

'--------------------------------------------------------------------------------
Private Sub cmdStart_Click()
'--------------------------------------------------------------------------------
' Controls the main transfer
'--------------------------------------------------------------------------------
' REVISIONS
' DPH 29/04/2002 - Set caption on cancel buton to be close
' DPH 01/05/2002 - Initialise overall transfer status
' REM 04/12/02 - added System Message transfer
' REM 03/04/03 - Test connection to server before trying data transfer
' REM 07/04/03 - Check to see if connection has failed or cancell has been clicked
'--------------------------------------------------------------------------------
Dim sErrorMsg As String
    
    On Error GoTo ErrHandler

    ' Disable start button
    cmdStart.Enabled = False
    cmdCancel.Caption = "&Cancel"
    
    ' DPH 01/05/2002 - Overall transfer status
    mbTransferOK = True
    
    'REM 07/04/03 - Connection status
    mbConnectionOK = True
    
    ' Set connection status setting
    Call SetConnectionStatus(csConnecting)
    Call SetTextBoxDisplay("Connecting to " & Me.MACROServerDesc & "...", trConnecting)

    'REM 03/04/03 - Test connection to server before doing data transfer
    If TestConnection(sErrorMsg) Then
        ' Get messages / data from server
        If DisplayGetMessages Then
            If Not mbCancelled Then
                ' Send data from the site to the server
                DisplaySendMessages
            Else
                GoTo Cancelled
            End If
        End If
        
        If Not mbConnectionOK Then GoTo ConnectFail
        If mbCancelled Then GoTo Cancelled
        
        ' DPH 24/01/2003 - Get modified reports from the server
        Call DisplayGetModifiedReports
        
        'REM 07/04/03 - Check to see if connection has failed or cancell has been clicked
        If Not mbConnectionOK Then GoTo ConnectFail
        If mbCancelled Then GoTo Cancelled
        
        'REM 04/12/02 - added system message transfer
        'send system messages from site to server
        If DisplaySendSystemMessages Then
        
            If Not mbCancelled Then
                'get system messages from server
                DisplayGetSystemMessages
            Else
                GoTo Cancelled
            End If
            
        End If
        
        'REM 07/04/03 - Check to see if connection has failed or cancell has been clicked
        If Not mbConnectionOK Then GoTo ConnectFail
        If mbCancelled Then GoTo Cancelled
        
        ' DPH 21/01/2005 - Added Pdu functionality
        ' check that PDU is enabled on this machine
        If LCase(GetMACROSetting(MACRO_SETTING_USE_PDU, "false")) = "true" Then
            ' timing loop for all of PDU section
            ' start timing
            Call mDataTransferTime.StartTiming(PDUFilesSection)

            ' send pdu messages from site to server
            DisplaySendPduMessages
            
            ' if not cancelled
            If Not mbCancelled Then
                ' get pdu messages from server
                DisplayGetPduMessages
            Else
                GoTo Cancelled
            End If
        
            ' stop timing
            Call mDataTransferTime.StopTiming(PDUFilesSection)

            ' Check to see if connection has failed or cancell has been clicked
            If Not mbConnectionOK Then GoTo ConnectFail
            If mbCancelled Then GoTo Cancelled
        
        End If
        
        ' DPH 24/01/2005 - close transfer
        'set progressbar to 100%
        Call SetProgressBar(100, "")
    
        Call SetConnectionStatus(csTransferComplete)
        ' set overall status of transfer
        Call SetTextBoxDisplay("", trTransferFinished, mbTransferOK)
        
        ' save timings for transfer to registry
        Call mDataTransferTime.SaveRegSettings
        
        ' DPH 29/04/2002 - Set caption on cancel button to be close
        cmdCancel.Caption = "&Close"
        WebBrowser.Document.Close
    
    Else 'if connection fails tell user and exit data transfer
    
        Call SetConnectionStatus(csError)
        Call SetTextBoxDisplay(sErrorMsg, trTransferError)
        cmdCancel.Caption = "&Close"
        WebBrowser.Document.Close
        GoTo ConnectFail
    End If
    
Exit Sub

ConnectFail:
    'REM 08/04/03 - log the connection failure
    Call gLog(gsCONNECT_FAIL, "Connection the server " & Me.MACROServerDesc & " failed")
Exit Sub

Cancelled:
    'REM 07/04/03 - If transfer cancelled then end data transfer
    Call SetConnectionStatus(csTransferCancelled)
    Call SetTextBoxDisplay("Data transfer was cancelled by user " & goUser.UserName, trTransferCancelled)
    WebBrowser.Document.Close
    Call gLog(gsCANCEL_TRANSFER, "Data transfer was cancelled by user " & goUser.UserName)
Exit Sub

ErrHandler:
    'REM 07/04/03 - if there is a general error then write it to the data transfer form log
    Call SetConnectionStatus(csError)
    Call SetTextBoxDisplay("There was an unexpected error, transfer cancelled" & vbCrLf & "Error Description: " & Err.Description & "Error Number: " & Err.Number, trTransferError)
    WebBrowser.Document.Close
    cmdCancel.Caption = "&Close"
End Sub

'--------------------------------------------------------------------------------
Private Sub Form_Load()
'--------------------------------------------------------------------------------
    
On Error GoTo ErrHandler
    
    Me.Icon = frmMenu.Icon
    
    'REM 07/04/03 - added web control to display datatransfer log
    WebBrowser.Stop
    WebBrowser.Document.Write ""
    WebBrowser.Document.Close
    
    ' initialise display
    Call Display(Me.Site, 100)
    
    ' initialise HTTP control
    Call InitialiseHTTPControl
    
    ' Get username & password
    msUser = frmMenu.gTrialOffice.User
    msPassword = frmMenu.gTrialOffice.Password
    
    ' DPH 15/04/2002 - Set Form Caption
    Me.Caption = "MACRO Data Transfer Communications"
'    txtMessage.Text = "You are about to connect to your MACRO server to transfer data." & _
'        vbCrLf & "This transfer may take some time." & vbCrLf & "Click Start Transfer to begin." & vbCrLf
    'REM 08/04/03 - set the start transfer text
    WebBrowser.Document.Write GetBaseText & _
                "<body style='font-family:verdana,arial,helvetica ;FONT-SIZE: 8pt;'> " & _
                "<font color='blue'><b>You are about to connect to your MACRO server to transfer data." & "<br>" & _
                "This transfer may take some time." & "<br>" & "Click 'Start Transfer' to begin</b></font>" & "<p>"
    WebBrowser.Document.body.Scroll = "auto"

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

'--------------------------------------------------------------------------------
Public Property Get Cancel() As Boolean
'--------------------------------------------------------------------------------
' read only property to return whether user has pressed cancel
'--------------------------------------------------------------------------------

    Cancel = mbCancelled

End Property

'--------------------------------------------------------------------------------
Private Sub cmdCancel_Click()
'--------------------------------------------------------------------------------
' REVISIONS
' DPH 29/04/2002 - Only unload form once transfer has completed
'--------------------------------------------------------------------------------

    mbCancelled = True
    
    If cmdCancel.Caption = "&Close" Then
        Unload Me
        Screen.MousePointer = vbDefault
    Else
        ' display cancelling message in display
        Call SetTextBoxDisplay("Attempting to cancel transfer...", trTransferData)
    End If
    
End Sub

'---------------------------------------------------------------------
Public Property Get MACROServerDesc() As String
'---------------------------------------------------------------------

    MACROServerDesc = msMACROServerDesc

End Property

'---------------------------------------------------------------------
Public Property Let MACROServerDesc(ByVal vNewValue As String)
'---------------------------------------------------------------------

    msMACROServerDesc = vNewValue

End Property

'---------------------------------------------------------------------
Public Property Get HTTPAddress() As String
'---------------------------------------------------------------------

    HTTPAddress = msHTTPAddress

End Property

'---------------------------------------------------------------------
Public Property Let HTTPAddress(ByVal vNewValue As String)
'---------------------------------------------------------------------
    
    msHTTPAddress = vNewValue

End Property

'---------------------------------------------------------------------
Public Property Get Site() As String
'---------------------------------------------------------------------

    Site = msSite

End Property

'---------------------------------------------------------------------
Public Property Let Site(ByVal vNewValue As String)
'---------------------------------------------------------------------

    msSite = vNewValue

End Property

'---------------------------------------------------------------------
Public Function DisplaySendSystemMessages() As Boolean
'---------------------------------------------------------------------
'29/11/02
'Function that sends all site messages to the server
'---------------------------------------------------------------------
Dim oSysDataXfer As SysDataXfer
Dim sMessageText As String
Dim vConfirmation As Variant
Dim sConfirmation As String
Dim sConIds As String
Dim sErrorMessage As String
Dim sHTTPAddress As String
Dim sHTTPData As String
Dim sSQL As String
Dim rsMsg As ADODB.Recordset
Dim sSystemMsgs As String
    
    On Error GoTo Errlabel

    Set oSysDataXfer = New SysDataXfer
    Call SetConnectionStatus(csSendingSystemDataToServer)
    Call SetTextBoxDisplay("Checking for system messages on the site " & Me.Site, trStepStarted)
    
    Call mDataTransferTime.StartTiming(SystemMessagesUp)
    
    DoEvents
    
    'check if there are any system messages on the site
    If oSysDataXfer.CheckSystemMessages(goUser.UserName, Me.Site, goUser.Database.DatabaseCode, sErrorMessage) Then
        Call SetTextBoxDisplay("There are system messages to send:-", trTransferData)
        
        'get a list of the messages to be sent
        sSQL = "SELECT MessageTimestamp,MessageBody FROM Message WHERE TrialSite = '" & Me.Site & "'" _
            & " AND MessageReceived = 0 AND MessageDirection = 1 AND MessageType >= 32" _
            & " ORDER BY MessageTimeStamp"
        Set rsMsg = New ADODB.Recordset
        rsMsg.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
        
        If mDataTransferTime.SystemMessagesUpTotal = 0 Then
            mDataTransferTime.SystemMessagesUpTotal = rsMsg.RecordCount
        End If
        
        Do While Not rsMsg.EOF
            sSystemMsgs = sSystemMsgs & rsMsg!MessageBody & "(" & CDate(rsMsg!MessageTimeStamp) & ")" & Chr(13) & Chr(10)
            rsMsg.MoveNext
        Loop
        rsMsg.Close
        Set rsMsg = Nothing
        
        Call SetTextBoxDisplay(sSystemMsgs, trTransferData)
        
        Call SetTextBoxDisplay("Sending to " & Me.MACROServerDesc & "...", trTransferData)
        
        'get the sites system messages
        sMessageText = oSysDataXfer.GetSystemMessages(Me.Site, goUser.Database.DatabaseCode, sErrorMessage)
        'If there is a message returned its because of an error
        If sErrorMessage <> "" Then GoTo Errlabel
        
        sHTTPAddress = Me.HTTPAddress & msWRITE_SYSTEM_MESSAGES_URL
        
        Do While sMessageText <> "."
            sErrorMessage = ""
            sHTTPData = "systemmessage=" & sMessageText
            
            On Error GoTo Timeout
            'send the messages to the server
            sConfirmation = PostDataHTTP(sHTTPAddress, mnDataTransferTimeout, sHTTPData)
            
            If sConfirmation = gsERRMSG_SEPARATOR Then
                sErrorMessage = "Error while sending system messages"
                GoTo Errlabel
            End If
            
            On Error GoTo Errlabel
            
            vConfirmation = Split(sConfirmation, gsERRMSG_SEPARATOR)
            
            sConIds = vConfirmation(0)
            sErrorMessage = vConfirmation(1)
            
            'error
            If sErrorMessage <> "" Then GoTo Errlabel
            
            
            'get the next system messages
            sMessageText = oSysDataXfer.GetSystemMessages(Me.Site, goUser.Database.DatabaseCode, sErrorMessage, , sConIds)
            'If there is a message returned its because of an error
            If sErrorMessage <> "" Then GoTo Errlabel
            DoEvents
        Loop
        Call SetTextBoxDisplay("System data has been sent to server " & Me.MACROServerDesc, trStepCompleted)

    Else
        'if there is an error message then end the transfer
        If sErrorMessage <> "" Then
            GoTo Errlabel
        Else 'else no messages to send
            Call SetConnectionStatus(csTransferComplete)
            Call SetTextBoxDisplay("No system messages on site " & Me.Site, trStepCompleted)
        End If
    End If
        
    ' Set Progress Bar to 85%
    Call SetProgressBar(85, "")
    Call mDataTransferTime.IncrementSystemMessagesUp
    Call mDataTransferTime.StopTiming(SystemMessagesUp)

    DisplaySendSystemMessages = True

Exit Function
Timeout:
    DisplaySendSystemMessages = False
    DisplayTimeoutMessage
    
Exit Function
Errlabel:
    If sErrorMessage = "" Then
        sErrorMessage = "Error while sending system messages. Error Description: " & Err.Description & ", Error Number: " & Err.Number
    End If
    Call SetTextBoxDisplay("System message send error, " & sErrorMessage, trStepCompleted, False)
    sErrorMessage = ReplaceInvalidCharsForDataXfer(sErrorMessage)
    Call gLog(gsSYSMSG_SEND_ERR, sErrorMessage)
    DisplaySendSystemMessages = True 'set to true so that DisplayGetSystemMessages still runs even if this fails
    mbTransferOK = False
End Function

'---------------------------------------------------------------------
Public Sub DisplayGetSystemMessages()
'---------------------------------------------------------------------
'REM 29/11/02
'Function that gets the System messages from the server and writes them to the site database
'---------------------------------------------------------------------
' REVISIONS
' DPH 24/01/2005 - System Messages finishes at 98% to allow for PDU
'---------------------------------------------------------------------
Dim sMessageText As String
Dim sHTTPAddress As String
Dim sHTTPData As String
Dim sMessageList As String
    
    On Error GoTo ErrHandler
    
    'Down load system messages
    sMessageText = ""
    
    Call SetConnectionStatus(csDownloadingSystemDataFromServer)
    Call SetTextBoxDisplay("Checking for system messages on server " & Me.MACROServerDesc & "...", trStepStarted)
    
    Call mDataTransferTime.StartTiming(SystemMessagesDown)
    
    DoEvents
    
    sHTTPAddress = Me.HTTPAddress & msCHECK_SYSTEM_MESSAGES_URL
    sHTTPData = "username=" & goUser.UserName & "&site=" & Me.Site
    
    sMessageText = PostDataHTTP(sHTTPAddress, mnDataTransferTimeout, sHTTPData)
    
    If sMessageText = "" Then
        GoTo Timeout
    ElseIf InStr(1, LCase(sMessageText), "<html") > 0 Then
        GoTo ASPError
    ElseIf sMessageText = "No system messages to download" Then
        Call SetTextBoxDisplay(sMessageText, trStepCompleted)
    ElseIf sMessageText = "There are system messages to download" Then
    
        Call SetTextBoxDisplay(sMessageText & ":-", trTransferData)
        
        sHTTPAddress = Me.HTTPAddress & msSYSTEM_MESSAGE_LIST_URL
        sHTTPData = "site=" & Me.Site
        
        'sometimes messages can be created during the transfer process and the MessageDown total could be 0, so check and if so just set it to 1
        If mDataTransferTime.SystemMessagesDownTotal = 0 Then
            mDataTransferTime.SystemMessagesDownTotal = 1
        End If
        
        'get list of messages to be downloaded from the server
        sMessageList = PostDataHTTP(sHTTPAddress, mnDataTransferTimeout, sHTTPData)
        'display the messages if there are no errors
        If (sMessageList <> "") And (InStr(1, LCase(sMessageText), "<html") = 0) Then
            Call SetTextBoxDisplay(sMessageList, trTransferData)
        Else
            GoTo ASPError
        End If
        
        'download the system messages
        Call DownLoadSystemMessages
        
    End If
    
    ' DPH 24/01/2005 - progress to 98% & finish transfer in cmdStart
    'set progressbar to 100%
    Call SetProgressBar(98, "")
    
    'Call SetConnectionStatus(csTransferComplete)
    ' DPH 01/05/2002 - set overall status of transfer
    'Call SetTextBoxDisplay("", trTransferFinished, mbTransferOK)
    
    ' save timings for transfer to registry
    'Call mDataTransferTime.SaveRegSettings


Exit Sub
Timeout:
    DisplayTimeoutMessage

Exit Sub
ASPError:
    Call SetConnectionStatus(csError)
    Call SetTextBoxDisplay("The MACRO server is not responding correctly. Please contact your system administrator.", trTransferError)
    'Call DialogError("The MACRO server is not responding correctly. Please contact your system administrator." _
        , "Data Transfer Call Error")
    mbConnectionOK = False
    cmdCancel.Caption = "&Close"
    
Exit Sub
ErrHandler:
  Call SetConnectionStatus(csError)
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "DisplayGetMessages", "frmDataTransfer")
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
Public Function ForgottenPassword(sSecurityCon As String, sDatabaseCode As String, sUsername As String, _
                                  sPassword As String, sSiteCode As String, sHTTPAddress As String, _
                                  sIISUserName As String, sIISPassword As String, _
                                  nPortNumber As Integer, sErrMsg As String) As eDTForgottenPassword
'---------------------------------------------------------------------
'REM 06/12/02
'Used to check and retrieve new user passwords from the server when a site user has forgotten their password
'and it has been chaned by the System Admin
'---------------------------------------------------------------------
Dim sSystemMessages As String
Dim sFullHTTPAddress As String
Dim sHTTPData As String
Dim oSysMessages As SysDataXfer
Dim vSystemMessages As Variant
Dim sSysMsg As String
Dim sConfirmationIds As String

    
    On Error GoTo Errlabel

    sSystemMessages = ""

    ' initialise HTTP control
    Call InitialiseHTTPControl
    
    sFullHTTPAddress = sHTTPAddress & msFORGOTTEN_PASSWORD_URL
    sHTTPData = "username=" & sUsername & "&password=" & sPassword & "&site=" & sSiteCode
    
    msUser = sIISUserName
    msPassword = sIISPassword
    
    sSystemMessages = PostDataHTTP(sFullHTTPAddress, mnDataTransferTimeout, sHTTPData, nPortNumber)
    
    vSystemMessages = Split(sSystemMessages, gsERRMSG_SEPARATOR)
    
    sSysMsg = vSystemMessages(0)
    sErrMsg = vSystemMessages(1)
    
    If sSysMsg = gsEND_OF_MESSAGES Then GoTo NoPassword
    
    'If there is a err message returned then exit the function and return the message
    If sErrMsg <> "" Then GoTo Errlabel
    
    Set oSysMessages = New SysDataXfer

    'write the system message to the database and to the Message table
    sConfirmationIds = oSysMessages.WriteNewPassword(sDatabaseCode, sSysMsg, sErrMsg)
    
    'If there is a err message returned then exit the function and return the message
    If sErrMsg <> "" Then GoTo Errlabel
    
    'add the confrimation ids to the address
    sHTTPData = "username=" & sUsername & "&site=" & sSiteCode & "&confirmationids=" & sConfirmationIds
    
    'return the confirmation Id's
    sSystemMessages = PostDataHTTP(sFullHTTPAddress, mnDataTransferTimeout, sHTTPData, nPortNumber)
    
    vSystemMessages = Split(sSystemMessages, gsERRMSG_SEPARATOR)

    sSysMsg = vSystemMessages(0)
    sErrMsg = vSystemMessages(1)
    
    'If there is a err message returned then exit the function and return the message
    If sErrMsg <> "" Then GoTo Errlabel
        
    ForgottenPassword = pSuccess
    
Exit Function
NoPassword:
    ForgottenPassword = pNoPassword

Exit Function
Errlabel:

    If sErrMsg = "" Then
        sErrMsg = "Error Description: " & Err.Description & ", Error Number: " & Err.Number
        ForgottenPassword = pError
    ElseIf sSysMsg = "Account Locked out" Then
        ForgottenPassword = pIncorrectPassword
    ElseIf sErrMsg = "Incorrect password" Then
        ForgottenPassword = pIncorrectPassword
    Else
        ForgottenPassword = pError
    End If
    
End Function

'---------------------------------------------------------------------
Public Function DisplayGetMessages() As Boolean
'---------------------------------------------------------------------
' REVISIONS
' DPH 01/05/2002 - Replaced GetStringResponseHTTP with PostDataHTTP calls
'---------------------------------------------------------------------
Dim sMessageText As String
' This gets written by the ASP
Const sNO_LF_MESSAGES = "No Lock/Freeze messages to download"

    ' initialise data transfer time class
    Set mDataTransferTime = New clsDataTransferTime
    
    'enable the cancel button
    Me.cmdCancel.Visible = True
    Me.cmdCancel.Caption = "&Cancel"
    mbCancelled = False
    
    'ATN 24/2/2000, DoEvents to ensure that button is fully displayed
    DoEvents
    
    On Error GoTo ErrHandler
        
    ' firstly collect download info to estimate transfer time
    If Not CollectInfoAboutDownload Then
        GoTo ASPError
    End If
    Call CollectInfoAboutUpload
    
    ' set DataTransferTime settings
    ' totals
    'REM 12/12/02 - added two new parameters for system messages
    'NCJ 20/12/02 - added two new parameters for Lock/Freeze messages
    mDataTransferTime.Init mlNoTransLockFreeze, mlNoTransStudyUpdate, mlNoTransLaboratoryDown, _
                        mlNoTransMIMessageDown, mlNoTransMIMessageUp, mlNoTransSubjects, _
                        mlNoTransLaboratoryUp, mlNoTransStudyStatus, _
                        mlNoSysMessagesDownload, mlNoSysMessagesUpload, _
                        mlNoTransLFMessageDown, mlNoTransLFMessageUp, _
                        mlNoTransReportFiles
    
    '   Check if cancel button has been pressed
    If mbCancelled Then
        DisplayGetMessages = False
        Exit Function
    End If

    On Error GoTo Timeout
    
    ' DPH 01/05/2002 - Changed call so throws error to be caught gracefully rather than MACRO error
    'Get study messages if there are any
    sMessageText = PostDataHTTP(Me.HTTPAddress & msMessageListURL & "?site=" & Me.Site, mnDataTransferTimeout)
        
    On Error GoTo ErrHandler
    
    ' Display message info (if not error)
    If sMessageText <> "" And InStr(1, LCase(sMessageText), "<html") = 0 Then
        Call SetTextBoxDisplay("Connected to server", trStepCompleted)
        Call SetConnectionStatus(csConnectionEstablished)
        Call SetConnectionStatus(csDownloadingFromServer)
        Call SetTextBoxDisplay(sMessageText, trStepStarted)
    End If
    
    If sMessageText = "" Then
        GoTo Timeout
    ElseIf InStr(1, LCase(sMessageText), "<html") > 0 Then
        GoTo ASPError
    ElseIf Left$(sMessageText, 29) <> "No study messages to download" Then
        Call SetConnectionStatus(csDownloadingFromServer)
        
        'if there is a connection error in this routine then exit
        If Not DownloadMessages Then GoTo Timeout
        
    End If
    
    ' Set Progress Bar to 10%
    Call SetProgressBar(10, "")
    
    '   Check if cancel button has been pressed
    If mbCancelled Then
        DisplayGetMessages = False
        Exit Function
    End If
    
    'Mo Morris 18/6/00, repeat the above for MIMessages
    sMessageText = ""
    
    On Error GoTo Timeout
    
    ' DPH 01/05/2002 - Changed call so throws error to be caught gracefully rather than MACRO error
    'Get any user messages
    sMessageText = PostDataHTTP(Me.HTTPAddress & msMIMessageListURL & "?site=" & Me.Site, mnDataTransferTimeout)
    
    On Error GoTo ErrHandler
    
    ' Display message info (if not error)
    If sMessageText <> "" And InStr(1, LCase(sMessageText), "<html") = 0 Then
        Call SetConnectionStatus(csDownloadingFromServer)
        Call SetTextBoxDisplay(sMessageText, trTransferData)
    End If
    
    If sMessageText = "" Then
        GoTo Timeout
    ElseIf InStr(1, LCase(sMessageText), "<html") > 0 Then
        GoTo ASPError
    ElseIf Left$(sMessageText, 28) <> "No user messages to download" Then
    
        If Not DownLoadMIMessages Then GoTo Timeout
        
    Else
        ' DPH 29/04/2002 - Not need change caption on cancel button
        'Me.cmdCancel.Caption = "&Close"
        Me.cmdCancel.Visible = True
    End If

    ' Set Progress Bar to 20%
    Call SetProgressBar(20, "")

    '   Check if cancel button has been pressed
    If mbCancelled Then
        'Call SetTextBoxDisplay("Transfer cancelled", trTransferCancelled)
        DisplayGetMessages = False
        Exit Function
    End If
    
    '************************************************************************
    'NCJ 20 Dec 02 - Repeat for LFMessages
    sMessageText = ""
    
    On Error GoTo Timeout
    'Get Lock/Freeze messages
    sMessageText = PostDataHTTP(Me.HTTPAddress & msLFMessageListURL & "?site=" & Me.Site, mnDataTransferTimeout)
    
    On Error GoTo ErrHandler
    
    ' Display message info (if not error)
    If sMessageText <> "" And InStr(1, LCase(sMessageText), "<html") = 0 Then
        Call SetConnectionStatus(csDownloadingFromServer)
        Call SetTextBoxDisplay(sMessageText, trTransferData)
    End If
    
    If sMessageText = "" Then
        GoTo Timeout
    ElseIf InStr(1, LCase(sMessageText), "<html") > 0 Then
        GoTo ASPError
    ElseIf Left$(sMessageText, Len(sNO_LF_MESSAGES)) <> sNO_LF_MESSAGES Then
    
        ' There are some LF messages
        If Not DownLoadLFMessages Then GoTo Timeout
        
    Else
        Me.cmdCancel.Visible = True
    End If

    ' NCJ 19 Dec 02 - Set Progress Bar to 25% (a guess!)
    Call SetProgressBar(25, "")

    '************************************************************************


    Call SetTextBoxDisplay("Completed download from " & Me.MACROServerDesc, trStepCompleted)

    If mbCancelled Then
        'Call SetTextBoxDisplay("Transfer cancelled", trTransferCancelled)
        DisplayGetMessages = False
        Exit Function
    End If
    
    DisplayGetMessages = True
    
Exit Function

Timeout:
    DisplayTimeoutMessage
    DisplayGetMessages = False
Exit Function

ASPError:
    Call SetConnectionStatus(csError)
    'REM 07/04/03 - changed error handling to write to data transfer form log
    Call SetTextBoxDisplay("The MACRO server is not responding correctly. Please contact your system administrator.", trTransferError)
    'Call DialogError("The MACRO server is not responding correctly. Please contact your system administrator." _
        , "Data Transfer Call Error")
    cmdCancel.Caption = "&Close"
    DisplayGetMessages = False
    mbConnectionOK = False
    mbTransferOK = False
Exit Function

ErrHandler:
  mbTransferOK = False
  Call SetTextBoxDisplay("An error was encountered while writing messages recieved from the server" & vbCrLf & "Error Description: " & Err.Description & "Error Number: " & Err.Number, trStepCompleted, mbTransferOK)
  DisplayGetMessages = True 'set to true as can still continue with the rest of data transfer

End Function

'---------------------------------------------------------------------
Public Sub DisplaySendMessages()
'---------------------------------------------------------------------
'Changed by Mo Morris 14/12/1999
'As per Macro 1.6 this sub now creates a recordset of all the patients that contain
'changed data and then loops through the recordset, calling AutoexportPRD to create
'a Cab Export file one patient at a time
'Changed Mo Morris 28/3/00 SR 3218
'Patient data export failing on files > 90kb. The solution entails the calling of three
'additional ASP script on the receiving server:-
'   exchange_receive_data.asp   (the original script, for less than 90kb)
'   exchange_receive_data1.asp  (for the first 90kb of a file greater than 90kb in total size)
'   exchange_receive_data2.asp  (for consequent blocks of 90kb of a file greater than 90kb in total size)
'   exchange_receive_data3.asp  (for the final block of less than 90kb in a file greater than 90kb in total size)
' DPH 14/1/2002 - Added delay if Inet control GetChunk call does not return anything from asp pages
' DPH 14/1/2002 - Changed because of problem where DOS window is launched giving user the opportunity
'   to save a changed eForm allowing data to go missing from the server as Changed flag is reset ...
' DPH 25/03/2002 - Remove Me.Show as now modal
'                   Added Integrity Check routine
'                   Validation of CAB / Checksum with server
' NCJ 23/24 Dec 02 - Added Lock Freeze message stuff
' REM 08/04/03 - if there is a connection error exit function
'---------------------------------------------------------------------
Dim vFilename As String
Dim msResult As String
Dim sSQL As String
Dim rsChangedRecords As ADODB.Recordset
Dim lCurrentClinicalTrialId As Long
Dim sCurrentTrialSite As String
Dim lCurrentPersonId As Long
Dim lTotalNumberToSend As Long
Dim lTotalSent As Long
'changed Mo Morris 27/3/00 Sr 3218
Dim nFileNumber As Integer
Dim sFullFileName As String
Dim sFileContent As String
'DPH 14/1/2002 - Fixes Data Integrity Problem + communication failure
Dim dTime As Double
Dim nTry As Integer
Dim alClinicalTrial() As Long
Dim asTrialSite() As String
Dim alPersonId() As Long
Dim lCount As Long
Dim sToken As String
' DPH 15/04/2002 - Add collection of files / filesizes for CAB validation
Dim colCABContents As New Collection
Dim sCABChecksum As String
Dim oMACROCheckSum As CheckSum
' 18/04/2002 - Added in http address control
Dim sHTTPAddress As String
Dim sHTTPData As String
Dim sTrialName As String
Dim lLastLFMessageId As Long
' Used to control percentage increment
Dim nPercentInc As Integer
' NCJ 24 Dec 02
Dim oLockFreezer As LockFreeze
Dim sMessage As String


    Set oMACROCheckSum = New CheckSum
    Set oLockFreezer = New LockFreeze
    
    Me.cmdCancel.Visible = True
    Me.cmdCancel.Caption = "&Cancel"
    
    Call SetConnectionStatus(csSendingDataToServer)
    Call SetTextBoxDisplay("", trStepStarted)
        
    '   ATN 24/2/2000
    '   DoEvents to ensure that button is fully displayed
    DoEvents
    
    Set oExchange = New clsExchange

    'create a recordset of all the subjects that contain changed data
    sSQL = "SELECT ClinicalTrialId, TrialSite,PersonId FROM TrialSubject WHERE Changed = " & Changed.Changed
    Set rsChangedRecords = New ADODB.Recordset
    rsChangedRecords.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    '   ATN 24/2/2000
    '   Get the initial total to send
    lTotalNumberToSend = rsChangedRecords.RecordCount
    lTotalSent = 0
    
    ' DPH 14/1/2002 - Collect & store recordset in the array (if not EOF)
    'If rsChangedRecords.RecordCount > 0 Then
    If Not rsChangedRecords.EOF Then
    
        Call SetConnectionStatus(csSendingDataToServer)
        Call SetTextBoxDisplay("Sending data to " & Me.MACROServerDesc & "...", trTransferData)
        
        ReDim alClinicalTrial(lTotalNumberToSend)
        ReDim asTrialSite(lTotalNumberToSend)
        ReDim alPersonId(lTotalNumberToSend)
        
        ' Store records in array
        lCount = 0
        
        'Do While Not rsChangedRecords.RecordCount = 0
        Do While Not rsChangedRecords.EOF
            
            ' Fill array
            alClinicalTrial(lCount) = rsChangedRecords!ClinicalTrialId
            asTrialSite(lCount) = rsChangedRecords!TrialSite
            alPersonId(lCount) = rsChangedRecords!PersonId
            
            lCount = lCount + 1
            rsChangedRecords.MoveNext
            
        Loop ' Do While Not rsChangedRecords.EOF
        rsChangedRecords.Close
        Set rsChangedRecords = Nothing
        
        Call SetTextBoxDisplay("Sending " & lTotalNumberToSend & " subjects ...", trTransferData)
        
        ' calculate percentage to increase each subject
        ' subject trans 40%
        nPercentInc = 1
        If lTotalNumberToSend > 0 Then
            nPercentInc = CInt(40 \ lTotalNumberToSend)
            If nPercentInc = 0 Then
                nPercentInc = 1
            End If
        End If
        ' start counting subjects
        Call mDataTransferTime.StartTiming(SubjectSection)
        
        For lCount = 0 To lTotalNumberToSend - 1
            
            DoEvents
            
            lCurrentClinicalTrialId = alClinicalTrial(lCount)
            sCurrentTrialSite = asTrialSite(lCount)
            lCurrentPersonId = alPersonId(lCount)
            
            '   DPH 14/1/2002
            '   Get lock on subject record if locked then ignore (on this transfer)
            sToken = oExchange.GetSubjectLock(goUser.UserName, lCurrentClinicalTrialId, sCurrentTrialSite, lCurrentPersonId)
            
            If Len(sToken) > 1 Then
            
                ' NCJ 23 Dec 02 - We need the Trialname and Last LFMessage ID for later
                sTrialName = TrialNameFromId(lCurrentClinicalTrialId)
                lLastLFMessageId = oLockFreezer.LastUsedLFMessageId(MacroADODBConnection, _
                                        lCurrentClinicalTrialId, sCurrentTrialSite, lCurrentPersonId)
            
                'changed by Mo Morris 14/12/99
                'note that transaction control covers the calls to AutoExportPRD
                'and UpdateChangedFlags for a single subject
                TransBegin
                            
                vFilename = oExchange.AutoExportPRD(Me.Site, lCurrentClinicalTrialId, sCurrentTrialSite, lCurrentPersonId, colCABContents)
    
                If vFilename > "" And GetFileLength(gsOUT_FOLDER_LOCATION & vFilename) > 0 Then
                    
                    On Error GoTo Timeout
                    
                    ' sending data message
                    Call SetTextBoxDisplay("Sending Data for Subject " & sCurrentTrialSite & "/" & lCurrentPersonId, trTransferData)
                    
                    'changed Mo Morris 9/2/00
'                    HEXEncodeFile gsOUT_FOLDER_LOCATION & vFilename, gsOUT_FOLDER_LOCATION & "_" & vFilename
                    HEXEncodeFileXCeed gsOUT_FOLDER_LOCATION & vFilename, gsOUT_FOLDER_LOCATION & "_" & vFilename
                    
                    ' DPH 25/03/2002 - Validate CAB File. If fails quit loop
                    If Not ValidateZIP(gsOUT_FOLDER_LOCATION & "_" & vFilename, colCABContents) Then
                        ' Log Error
                        gLog gsVALIDATE_ZIP, "CAB " & vFilename & " failed to validate correctly."
                        ' Remove subject lock
                        Call oExchange.RemoveSubjectLock(lCurrentClinicalTrialId, sCurrentTrialSite, lCurrentPersonId, sToken)
                        'ZA 24/06/2002 - set mbTransferOK flag to false
                        mbTransferOK = False
                        ' message on screen
                        Call SetTextBoxDisplay("Data not sent for Subject " & sCurrentTrialSite & "/" & lCurrentPersonId & " as validation of data file failed", trTransferData)
                    Else
                        ' Checksum Hex CAB & send
                        sCABChecksum = oMACROCheckSum.GetFileCheckSum(gsOUT_FOLDER_LOCATION & "_" & vFilename)
                        
                        nFileNumber = FreeFile
                        sFullFileName = gsOUT_FOLDER_LOCATION & "_" & vFilename
                        Open sFullFileName For Input As nFileNumber
                        sFileContent = Input(FileLen(sFullFileName), #nFileNumber)
                        If Len(sFileContent) < 90000 Then
                                
                            ' Send via new Web control
                            sHTTPAddress = Me.HTTPAddress & msReceiveDataURL
                            sHTTPData = "TrialOffice=" & _
                                Me.MACROServerDesc & "&Site=" & Me.Site & "&FileName=" & vFilename & "&Hex=" & _
                                sFileContent & "&chksum=" & sCABChecksum
                                ' NCJ 23 Dec 02 - Include Study name and Subject and LastLFMessageID
                                sHTTPData = sHTTPData & "&TrialName=" & sTrialName _
                                        & "&SubjectId=" & lCurrentPersonId & "&LastLFMessageId=" & lLastLFMessageId
                            ' get result
                            msResult = PostDataHTTP(sHTTPAddress, 300, sHTTPData)
                        Else
                                
                            sHTTPAddress = Me.HTTPAddress & msReceiveData1URL
                            sHTTPData = "TrialOffice=" & _
                                Me.MACROServerDesc & "&Site=" & Me.Site & "&FileName=" & vFilename & "&Hex=" & _
                                Left$(sFileContent, 90000)
                            ' get result
                            msResult = PostDataHTTP(sHTTPAddress, 300, sHTTPData)
                                
                            sFileContent = Mid$(sFileContent, 90001)
                            While Len(sFileContent) > 90000
                                
                                sHTTPAddress = Me.HTTPAddress & msReceiveData2URL
                                sHTTPData = "TrialOffice=" & _
                                    Me.MACROServerDesc & "&Site=" & Me.Site & "&FileName=" & vFilename & "&Hex=" & _
                                    Left$(sFileContent, 90000)
                                ' get result
                                msResult = PostDataHTTP(sHTTPAddress, 300, sHTTPData)
                                                                
                                sFileContent = Mid$(sFileContent, 90001)
                            Wend
                        
                            sHTTPAddress = Me.HTTPAddress & msReceiveData3URL
                            sHTTPData = "TrialOffice=" & _
                                Me.MACROServerDesc & "&Site=" & Me.Site & "&FileName=" & vFilename & "&Hex=" & _
                                Left$(sFileContent, 90000) & "&chksum=" & sCABChecksum
                            ' NCJ 23 Dec 02 - Include Study name and Subject and LastLFMessageID
                            sHTTPData = sHTTPData & "&TrialName=" & sTrialName _
                                        & "&SubjectId=" & lCurrentPersonId & "&LastLFMessageId=" & lLastLFMessageId
                            ' get result
                            msResult = PostDataHTTP(sHTTPAddress, 300, sHTTPData)
                        End If
                        Close nFileNumber
            
                        DoEvents

                        Select Case msResult
                        Case msASP_RETURN_SUCCESS
                        
                            sMessage = "Subject " & sCurrentTrialSite & "/" & lCurrentPersonId & " data file _" & vFilename & " sent to " & Me.MACROServerDesc & " successfully"
                            ' write message to screen
                            Call SetTextBoxDisplay(sMessage, trTransferData)
                            
                            '   On success, update the changed flags to indicate that the data has been sent
                            oExchange.UpdateChangedFlags lCurrentClinicalTrialId, sCurrentTrialSite, lCurrentPersonId
                            '   Captions will be updated after all records transferred
                            '   Increment counter of number already sent through
                            lTotalSent = lTotalSent + 1
                            
                            'REM 01/09/03 - added log success
                            gLog gsPATDATA_SEND, sMessage
                            
                            'REM 01/09/03 - add to Message table so that keeps a history of all sent messages (this message only used by Transfer History)
                            Call InsertMessage(sCurrentTrialSite, lCurrentClinicalTrialId, ExchangeMessageType.PatientDataSent, SQLStandardNow, 0, goUser.UserName, sMessage, "", 0, MessageReceived.Received)
                            
                        Case msASP_RETURN_CHECKSUM_ERROR
                        
                            sMessage = "Error in Checksum for Subject " & sCurrentTrialSite & "/" & lCurrentPersonId & " data file " & "_" & vFilename & " sending to " & Me.MACROServerDesc
                            ' write message to screen
                            Call SetTextBoxDisplay(sMessage, trTransferData)
                            
                            '   Increment counter of number already sent through
                            lTotalSent = lTotalSent + 1
                    
                            ' Log error
                            gLog gsPATDATA_SEND, sMessage
                            
                            mbTransferOK = False
                        Case Else
                            sMessage = "An unexpected error occurred. No more data could be sent to " & Me.MACROServerDesc
                            ' write message to screen
                            'ZA 24/06/2002 - changed message from "Connection broken/lost" to
                            '"unexpected error"
                            Call SetTextBoxDisplay(sMessage, trTransferData)
                            
                            Call oExchange.RemoveSubjectLock(lCurrentClinicalTrialId, sCurrentTrialSite, lCurrentPersonId, sToken)
                            'changed Mo Morris 21/3/00, If transfer fails there should be a transaction rollback
                            TransRollBack
                            'exit do
                            mbTransferOK = False
                            
                            'REM 01/09/03 - added Log error
                            gLog gsPATDATA_SEND, sMessage
                            
                            Exit For
                        End Select
                        
                        TransCommit
                        
                        ' DPH 14/1/2002 - unlock record
                        Call oExchange.RemoveSubjectLock(lCurrentClinicalTrialId, sCurrentTrialSite, lCurrentPersonId, sToken)
    
                        '   Check if cancel button has been pressed
                        If mbCancelled Then
                            'Call SetTextBoxDisplay("Transfer cancelled", trTransferCancelled)
                            Exit Sub
                        End If
                        
                    End If
                Else
                    ' DPH 14/1/2002 - unlock subject if no file to open
                    Call oExchange.RemoveSubjectLock(lCurrentClinicalTrialId, sCurrentTrialSite, lCurrentPersonId, sToken)
                    ' Show Failed message
                    Call SetTextBoxDisplay("Subject " & sCurrentTrialSite & "/" & lCurrentPersonId & " has not been sent as failed to create data file.", trTransferData)
                    mbTransferOK = False
                End If
                
                ' DPH 15/04/2002 - Make sure no transactions still running
                If gnTransactionControlOn > 0 Then
                    ' if a transaction still open then check has failed so commit log entries
                    TransCommit
                End If
            Else
                ' Show locked record
                Call SetTextBoxDisplay("Subject " & sCurrentTrialSite & "/" & lCurrentPersonId & " is currently MACRO locked.", trTransferData)
            End If
            'Loop    'on rsChangedRecords.RecordCount
            ' update timer class
            Call mDataTransferTime.IncrementSubject
            ' Increment progress bar
            If (prgProgress.Value + nPercentInc) < 45 Then
                Call SetProgressBar(prgProgress.Value + nPercentInc, "")
            End If
        Next ' lCount
        
        ' Update captions now
        'ZA 24/06/2002 - passed mbTransferOK parameter
        Call SetTextBoxDisplay(lTotalSent & " record(s) transferred", trStepCompleted, mbTransferOK)
    Else
        Call SetTextBoxDisplay("No data to be sent to " & Me.MACROServerDesc, trTransferData)
    End If
    
    ' Set Progress Bar to 45%
    Call SetProgressBar(45, "")
    
    ' DPH 29/04/2002 - Check if cancel button has been pressed
    If mbCancelled Then
        Exit Sub
    End If
    
    'Changed Mo Morris 9/11/00, call new procedure SendLaboratory
    ' REM 08/04/03 - if there is a connection error in this function then exit current routine
    If Not SendLaboratory Then Exit Sub
    
    ' Set Progress Bar to 52%
    Call SetProgressBar(52, "")
    
    ' DPH 29/04/2002 - Check if cancel button has been pressed
    If mbCancelled Then
        'Call SetTextBoxDisplay("Transfer cancelled", trTransferCancelled)
        Exit Sub
    End If
    
    'Changed Mo Morris 22/5/00 call new procedure SendMIMessages
    ' REM 08/04/03 - if there is a connection error in this function then exit current routine
    If Not SendMIMessages Then Exit Sub
    
    ' Set Progress Bar to 64%
    Call SetProgressBar(64, "")
    
    ' DPH 29/04/2002 - Check if cancel button has been pressed
    If mbCancelled Then
        'Call SetTextBoxDisplay("Transfer cancelled", trTransferCancelled)
        Exit Sub
    End If
    
    ' NCJ 20 Dec 02 - SendLFMessages
    ' REM 08/04/03 - if there is a connection error in this function then exit current routine
    If Not SendLFMessages Then Exit Sub
    
    ' Set Progress Bar to 68% (a guess!)
    Call SetProgressBar(68, "")
    
    ' DPH 15/04/2002 - Added Integrity Check routine
    ' REM 08/04/03 - if there is a connection error in this function then exit current routine
    If Not PerformIntegrityCheck Then Exit Sub
    
    ' Set Progress Bar to 70%
    Call SetProgressBar(70, "")
    
    'REM 08/04/03 - Set status for completion of sending data to server
    Call SetTextBoxDisplay("Completed sending data to " & Me.MACROServerDesc, trStepCompleted)
    
    ' save timings for transfer to registry
'    Call mDataTransferTime.SaveRegSettings
    
    Screen.MousePointer = vbDefault
    
    Set oExchange = Nothing
    Set oMACROCheckSum = Nothing
    Set oLockFreezer = Nothing

Exit Sub

Timeout:

    'changed Mo Morris 21/3/00, check for the need for a transaction rollback
    If gnTransactionControlOn > 0 Then
        TransRollBack
        ' DPH 14/1/2002 - unlock subject
        Call oExchange.RemoveSubjectLock(lCurrentClinicalTrialId, sCurrentTrialSite, lCurrentPersonId, sToken)
    End If
    
    DisplayTimeoutMessage

End Sub

'---------------------------------------------------------------------
Private Sub DownLoadSystemMessages()
'---------------------------------------------------------------------
'REM 25/11/02
'returns any system messages from the server
'---------------------------------------------------------------------
Dim sSystemMessages As String
Dim oSysMessages As SysDataXfer
Dim sConfirmationIds As String
Dim vSystemMessages As Variant
Dim sHTTPAddress As String
Dim sHTTPData As String
Dim sSysMsg As String
Dim sErrMsg As String

    On Error GoTo Errlabel

    cmdCancel.Visible = True

    sHTTPAddress = Me.HTTPAddress & msGET_SYSTEM_MESSAGES_URL
    sHTTPData = "username=" & goUser.UserName & "&site=" & Me.Site
    
    Call SetTextBoxDisplay("Downloading system messages from " & Me.MACROServerDesc & "...", trTransferData)
    
    'get system messages from the server
    sSystemMessages = PostDataHTTP(sHTTPAddress, mnDataTransferTimeout, sHTTPData)
    
    'split the system messages from the possible error message return
    vSystemMessages = Split(sSystemMessages, gsERRMSG_SEPARATOR)
    
    sSysMsg = vSystemMessages(0)
    sErrMsg = vSystemMessages(1)

    'Error
    If sErrMsg <> "" Then GoTo Errlabel
    
    Set oSysMessages = New SysDataXfer
    
    'keep getting messages until return is "."
    Do While sSysMsg <> "."
        DoEvents
        'write the system message to the database and to the Message table
        sConfirmationIds = oSysMessages.WriteSystemMessage(goUser.Database.DatabaseCode, sSysMsg, sErrMsg)
        
        'If there is a message returned its because of an error
        If sErrMsg <> "" Then GoTo Errlabel
        
        If sConfirmationIds <> "" Then
            sHTTPData = "username=" & goUser.UserName & "&site=" & Me.Site & "&confirmationids=" & sConfirmationIds
            'get the next system messages from th eserver and send last messages confirmation Id's
            sSystemMessages = PostDataHTTP(sHTTPAddress, mnDataTransferTimeout, sHTTPData)
            
            vSystemMessages = Split(sSystemMessages, gsERRMSG_SEPARATOR)
        
            sSysMsg = vSystemMessages(0)
            sErrMsg = vSystemMessages(1)
            
            'If there is a message returned its because of an error
            If sErrMsg <> "" Then GoTo Errlabel
        End If
        
    Loop
    'indicate that the step is completed
    Call SetTextBoxDisplay("All system messages from " & Me.MACROServerDesc & " downloaded", trStepCompleted)
    
    Call mDataTransferTime.IncrementSystemMessagesDown
    Call mDataTransferTime.StopTiming(SystemMessagesDown)

    Set oSysMessages = Nothing
    
Exit Sub
Errlabel:
    If sErrMsg = "" Then
        sErrMsg = "Error Description: " & Err.Description & ", Error Number: " & Err.Number
    End If
    Call SetTextBoxDisplay("System message download error, " & sErrMsg, trStepCompleted, False)
    Call gLog(gsSYSMSG_DOWNLOAD_ERR, sErrMsg)
    mbTransferOK = False
End Sub

'---------------------------------------------------------------------
Private Function DownloadMessages() As Boolean
'---------------------------------------------------------------------
' REVISIONS
'---------------------------------------------------------------------
' DPH 29/04/2002 - Tightened error handling around study def / lab download
' DPH 01/05/2002 - Replaced GetStringResponseHTTP with PostDataHTTP calls
'               Improved error handling on handling response data
' DPH 02/05/2002 - Changed GetByteResponseHTTP function so error handled gracefully
' DPH 11/07/2002 - CBBL 2.2.19.31 & 2.2.19.34
'                Closing open recordset that caused transaction problems in SQL Server
' ATO 23/08/2002 - Added RepeatNumber to DownloadMessages.
' DPH 27/08/2002 - Handling change to data file names for study versioning
'                   New fields handled for Study Versioning
'---------------------------------------------------------------------
Dim msNextMessage As String
Dim msMessageId As String
Dim sTrialSite As String
Dim sClinicalTrialId As String
Dim msMessageType As String
Dim msMessageParameters As String
Dim sSQL As String
Dim rsSite As ADODB.Recordset
Dim rsTrialSite As ADODB.Recordset
Dim rsClinicalTrial As ADODB.Recordset
Dim rsLabSites As ADODB.Recordset
Dim msSDDFile As String
Dim msLDDFile As String
Dim mbSDD() As Byte
Dim mnFreeFile As Integer
Dim sCabImportFile As String
Dim sClinicalTrialName As String

Dim nLockSetting As LockStatus
Dim lPersonId As Long
Dim sTimestamp As String
Dim lVisitId As Long
Dim nVisitCycleNumber As Integer
Dim lCRFPageTaskId As Long
Dim lResponseTaskId As Long
Dim sLaboratory As String
'ATO 21/08/2002
Dim nRepeatNumber As Integer
'REM 17/01/03 - New parameters for creating a site
Dim vParameters As Variant
Dim sSiteParameters As String
Dim vSite As Variant
Dim sCABfile As String
Dim sSiteDescript As String
Dim nSiteStatus As Integer
Dim nSiteLocation As Integer
Dim nSiteCountry As Integer


' Used to control percentage increment
Dim nPercentInc As Integer
Dim bImportOK As Boolean

'used to store the split lockstatus message parameter
Dim vLockStatus As Variant

    On Error GoTo ErrHandler

    cmdCancel.Visible = True
    Call SetConnectionStatus(csDownloadingFromServer)
    'Call SetTextBoxDisplay("Downloading messages from " & Me.MACROServerDesc, trTransferData)
    nPercentInc = 1
    
    On Error GoTo Timeout
    
    ' DPH 01/05/2002 - Changed call so throws error to be caught gracefully rather than MACRO error
    'msNextMessage = GetStringResponseHTTP(Me.HTTPAddress & msGetNextMessageURL & "?site=" & Me.Site, mnDataTransferTimeout)
    msNextMessage = PostDataHTTP(Me.HTTPAddress & msGetNextMessageURL & "?site=" & Me.Site, mnDataTransferTimeout)
    
    On Error GoTo ErrHandler

    ' DPH 01/05/2002 - Pick up if an error has occurred in the ASP page
    If Left(msNextMessage, 5) = "ERROR" Or InStr(1, LCase(msNextMessage), "<html") > 0 Then
        Call SetTextBoxDisplay("Messages could not be downloaded", trStepCompleted, False)
        ' log error
        gLog gsDOWNLOAD_MESG, "Messages could not be downloaded."
        mbTransferOK = False
        Exit Function
    End If

    Do While msNextMessage <> "."
        msMessageId = ExtractFirstItemFromList(msNextMessage, "<br>")
        sTrialSite = ExtractFirstItemFromList(msNextMessage, "<br>")
        sClinicalTrialId = ExtractFirstItemFromList(msNextMessage, "<br>")
        msMessageType = ExtractFirstItemFromList(msNextMessage, "<br>")
        msMessageParameters = ExtractFirstItemFromList(msNextMessage, "<br>")
        DoEvents
        
        Select Case msMessageType
        Case ExchangeMessageType.NewTrial, ExchangeMessageType.NewVersion
            On Error GoTo Timeout
            
            'REM 17/01/03 - MessageParameters field now has CabFile and Site details for NewTrial Message type
            If msMessageType = ExchangeMessageType.NewTrial Then
                vParameters = Split(msMessageParameters, gsMSGSEPARATOR)
                msMessageParameters = vParameters(0)
                sSiteParameters = vParameters(1)
                
                'add all the site parameters to the site table
                If sSiteParameters <> "" Then
                    'does the site code exist
                    sSQL = "SELECT Site FROM Site WHERE Site = '" & sTrialSite & "'"
                    Set rsSite = New ADODB.Recordset
                    rsSite.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
                    'if site does not exist then insert it
                    If rsSite.EOF Then
                        'get all site parameters
                        vSite = Split(sSiteParameters, gsSEPARATOR)
                        sSiteDescript = vSite(1)
                        nSiteStatus = vSite(2)
                        nSiteLocation = vSite(3)
                        nSiteCountry = vSite(4)

                        'Insert new site
                        sSQL = "INSERT into Site (Site,SiteDescription,SiteStatus,SiteLocation,SiteCountry) " _
                            & " VALUES ('" & sTrialSite & "','" & sSiteDescript & "'," & nSiteStatus & "," & nSiteLocation & "," _
                            & nSiteCountry & ")"
                        MacroADODBConnection.Execute sSQL
                    End If
                    
                    rsSite.Close
                    Set rsSite = Nothing
                    
                End If
            End If
            
            'with NewTrial and NewVersion messages msMessageParameters contains the cab file that
            'needs to be downloaded, but if the Newtrial message was created before a Cab file existed
            'then msMessageParameter will be blank and the whole message should be ignored
            'Changed Mo Morris 28/4/00, SR3330, rTrim added to remove space that is added by SQLServer to emtpy strings
            If RTrim(msMessageParameters) <> "" Then
                ' start timing
                Call mDataTransferTime.StartTiming(StudyUpdateSection)
                ' set importok
                bImportOK = True
                'extract the TrialName from the cab file name
                ' DPH 27/08/2002 - Handling change to data file names for study versioning
                ' versioned files will have numeric prefix - strip these out
'                sClinicalTrialName = Mid(msMessageParameters, 1, InStr(msMessageParameters, ".") - 1)
                sClinicalTrialName = GetStudyNameFromParameterField(msMessageParameters)
                Call SetTextBoxDisplay("Downloading " & sClinicalTrialName & " study definition from " & Me.MACROServerDesc, trTransferData)
                sCabImportFile = msMessageParameters
                'call GetByteResponseHTTP which places contents of message into mbSDD
                ' DPH 02/05/2002 - Changed function so error handled gracefully
                'mbSDD() = GetByteResponseHTTP(Me.HTTPAddress & sCabImportFile, 300)
                mbSDD() = GetDataHTTP(Me.HTTPAddress & sCabImportFile, 300)
                mnFreeFile = FreeFile
                Open gsIN_FOLDER_LOCATION & sCabImportFile For Binary Access Write As #mnFreeFile
                Put #mnFreeFile, , mbSDD()
                Close #mnFreeFile
                
                On Error GoTo ErrHandler
                
                Set oExchange = New clsExchange
                
                'changed Mo Morris 23/2/00
                'Unpack the CAB file into an SDD file, and individual study documents and
                'graphic files and place them in directory AppPath/CabExtract
                oExchange.ImportStudyDefinitionCAB gsIN_FOLDER_LOCATION & sCabImportFile
            
                'changed Mo Morris 23/2/00
                msSDDFile = Dir(gsCAB_EXTRACT_LOCATION & "*.sdd")
            
                If msSDDFile > "" Then
                    'Import the extracted SDD file into current database
                    'changed Mo Morris 9/2/00
                    If oExchange.ImportSDD(gsCAB_EXTRACT_LOCATION & msSDDFile) = ExchangeError.Success Then
                                            
                        'Get the ClinicalTrialId for the newly imported trial
                        sSQL = "SELECT ClinicalTrialId FROM ClinicalTrial " _
                                & "WHERE ClinicalTrial.ClinicalTrialName = '" & sClinicalTrialName & "'"
                        Set rsClinicalTrial = New ADODB.Recordset
                        rsClinicalTrial.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
                        
                        'REM 17/01/03 - Moved this code above and added new site parameters for MACRO 3.0
'                        'Changed Mo Morris 28/4/00, SR3164, Add Site to table Site as well as table TrialSite
'                        'does the site code exist
'                        sSQL = "SELECT Site FROM Site WHERE Site = '" & sTrialSite & "'"
'                        Set rsSite = New ADODB.Recordset
'                        rsSite.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'                        'if site does not exist then insert it
'                        If rsSite.EOF Then
'                            ' DPH 27/08/2002 - SiteLocation for study versioning
'                            sSQL = "INSERT into Site (Site,SiteDescription,SiteStatus,SiteLocation) " _
'                                & " VALUES ('" & sTrialSite & "','" & sTrialSite & "',0,-1)"
'                            MacroADODBConnection.Execute sSQL
'                        End If

                        'does a trialsite exist for this trial
                        sSQL = "SELECT ClinicalTrialId FROM TrialSite " _
                                & "WHERE ClinicalTrialid = " & rsClinicalTrial!ClinicalTrialId
                        Set rsTrialSite = New ADODB.Recordset
                        rsTrialSite.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

                        'if no trialsite exists for this trial then insert one
                        If rsTrialSite.EOF Then
                            ' DPH 27/08/2002 - StudyVersion for study versioning
                            sSQL = "INSERT INTO TrialSite (ClinicalTrialid,TrialSite,StudyVersion) " _
                                & " VALUES (" & rsClinicalTrial!ClinicalTrialId & ",'" & sTrialSite & "',0)"

                            MacroADODBConnection.Execute sSQL
                        End If

                        rsClinicalTrial.Close
                        Set rsClinicalTrial = Nothing
                        rsTrialSite.Close
                        Set rsTrialSite = Nothing
                        ' DPH 11/07/2002 - CBBL 2.2.19.31 & 2.2.19.34 Closing open recordset that caused
                        '   transaction problems in SQL Server

                    Else
                        ' import failed
                        bImportOK = False
                    End If
                        
                Else
                    ' No data to import for study
                    bImportOK = False
                End If
                
                Set oExchange = Nothing

                'Kill the Cab File
                'changed Mo Morris 9/2/00
                Kill gsIN_FOLDER_LOCATION & sCabImportFile
                
                ' stop timing
                Call mDataTransferTime.IncrementStudyUpdate
                Call mDataTransferTime.StopTiming(StudyUpdateSection)
                
                If bImportOK Then
                    Call SetTextBoxDisplay(sClinicalTrialName & " study definition downloaded", trTransferData)
                Else
                    Call SetTextBoxDisplay(sClinicalTrialName & " study definition failed to download. Please contact your systems administrator", trTransferData)
                    ' Log Error
                    gLog gsDOWNLOAD_MESG, sClinicalTrialName & " study definition failed to download. Please contact your systems administrator"
                    mbTransferOK = False
                End If
            End If
        
        'Mo Morris 16/5/00, SR 3422, ExchangeMessageType.InPreparation added
        Case ExchangeMessageType.ClosedFollowUp, ExchangeMessageType.ClosedRecruitment, _
                ExchangeMessageType.TrialOpen, ExchangeMessageType.TrialSuspended, ExchangeMessageType.InPreparation
            ' start timing
            Call mDataTransferTime.StartTiming(StudyStatusSection)
            'Changed Mo Morris 4/5/00 SR3406
            sClinicalTrialName = Mid(msMessageParameters, 1, InStr(msMessageParameters, "*") - 1)
            msMessageParameters = Mid(msMessageParameters, InStr(msMessageParameters, "*") + 1)
            Call gdsUpdateTrialStatus(TrialIdFromName(sClinicalTrialName), gnCurrentVersionId(TrialIdFromName(sClinicalTrialName)), CInt(msMessageParameters))
            'stop timing
            Call mDataTransferTime.IncrementStudyStatus
            Call mDataTransferTime.StopTiming(StudyStatusSection)
        'Mo Morris 25/4/00 Lock/Unlock/freeze meassage processing added
        Case ExchangeMessageType.TrialSubjectLockStatus
            'Changed Mo Morris 4/5/00 SR3406
            'start timing
            Call mDataTransferTime.StartTiming(LockFreezeSection)
            'put message parameters into arrays
            vLockStatus = Split(msMessageParameters, "*")
            sClinicalTrialName = vLockStatus(0)
            nLockSetting = vLockStatus(1)
            lPersonId = vLockStatus(2)
            sTimestamp = vLockStatus(3)
            Call RemoteSetTrialSubjectLockStatus(TrialIdFromName(sClinicalTrialName), sTrialSite, lPersonId, nLockSetting, sTimestamp)
            'stop timing
            Call mDataTransferTime.IncrementLockFreeze
            Call mDataTransferTime.StopTiming(LockFreezeSection)
        Case ExchangeMessageType.VisitInstanceLockStatus
            'Changed Mo Morris 4/5/00 SR3406
            'start timing
            Call mDataTransferTime.StartTiming(LockFreezeSection)
            'put message parameters into arrays
            vLockStatus = Split(msMessageParameters, "*")
            sClinicalTrialName = vLockStatus(0)
            nLockSetting = vLockStatus(1)
            lPersonId = vLockStatus(2)
            lVisitId = vLockStatus(3)
            nVisitCycleNumber = vLockStatus(4)
            sTimestamp = vLockStatus(5)
            Call RemoteSetVisitInstanceLockStatus(TrialIdFromName(sClinicalTrialName), sTrialSite, lPersonId, lVisitId, nVisitCycleNumber, nLockSetting, sTimestamp)
            'stop timing
            Call mDataTransferTime.IncrementLockFreeze
            Call mDataTransferTime.StopTiming(LockFreezeSection)
        Case ExchangeMessageType.CRFPageInstanceLockStatus
            'Changed Mo Morris 4/5/00 SR3406
            'start timing
            Call mDataTransferTime.StartTiming(LockFreezeSection)
            'put message parameters into arrays
            vLockStatus = Split(msMessageParameters, "*")
            sClinicalTrialName = vLockStatus(0)
            nLockSetting = vLockStatus(1)
            lPersonId = vLockStatus(2)
            lCRFPageTaskId = vLockStatus(3)
            sTimestamp = vLockStatus(4)
            Call RemoteSetCRFPageInstanceLockStatus(TrialIdFromName(sClinicalTrialName), sTrialSite, lPersonId, lCRFPageTaskId, nLockSetting, sTimestamp)
            'stop timing
            Call mDataTransferTime.IncrementLockFreeze
            Call mDataTransferTime.StopTiming(LockFreezeSection)
        Case ExchangeMessageType.DataItemLockStatus
            'Changed Mo Morris 4/5/00 SR3406
            'start timing
            Call mDataTransferTime.StartTiming(LockFreezeSection)
            'put message parameters into arrays
            vLockStatus = Split(msMessageParameters, "*")
            sClinicalTrialName = vLockStatus(0)
            nLockSetting = vLockStatus(1)
            lPersonId = vLockStatus(2)
            lResponseTaskId = vLockStatus(3)
            nRepeatNumber = vLockStatus(4)
            sTimestamp = vLockStatus(5)
            Call RemoteSetDataItemLockStatus(TrialIdFromName(sClinicalTrialName), sTrialSite, lPersonId, lResponseTaskId, nLockSetting, nRepeatNumber, sTimestamp)
            'stop timing
            Call mDataTransferTime.IncrementLockFreeze
            Call mDataTransferTime.StopTiming(LockFreezeSection)
        Case ExchangeMessageType.TrialSubjectUnLock
            'Changed Mo Morris 4/5/00 SR3406
            'start timing
            Call mDataTransferTime.StartTiming(LockFreezeSection)
            'put message parameters into arrays
            vLockStatus = Split(msMessageParameters, "*")
            sClinicalTrialName = vLockStatus(0)
            lPersonId = vLockStatus(1)
            Call RemoteUnlockTrialSubject(TrialIdFromName(sClinicalTrialName), sTrialSite, lPersonId)
            'stop timing
            Call mDataTransferTime.IncrementLockFreeze
            Call mDataTransferTime.StopTiming(LockFreezeSection)
        Case ExchangeMessageType.VisitInstanceUnLock
            'Changed Mo Morris 4/5/00 SR3406
            'start timing
            Call mDataTransferTime.StartTiming(LockFreezeSection)
            'put message parameters into arrays
            vLockStatus = Split(msMessageParameters, "*")
            sClinicalTrialName = vLockStatus(0)
            lPersonId = vLockStatus(1)
            lVisitId = vLockStatus(2)
            nVisitCycleNumber = vLockStatus(3)
            Call RemoteUnlockVisitInstance(TrialIdFromName(sClinicalTrialName), sTrialSite, lPersonId, lVisitId, nVisitCycleNumber)
            'stop timing
            Call mDataTransferTime.IncrementLockFreeze
            Call mDataTransferTime.StopTiming(LockFreezeSection)
        Case ExchangeMessageType.CRFPageInstanceUnLock
            'start timing
            Call mDataTransferTime.StartTiming(LockFreezeSection)
            'put message parameters into arrays
            vLockStatus = Split(msMessageParameters, "*")
            'Changed Mo Morris 4/5/00 SR3406
            sClinicalTrialName = vLockStatus(0)
            lPersonId = vLockStatus(1)
            lCRFPageTaskId = vLockStatus(2)
            lVisitId = vLockStatus(3)
            nVisitCycleNumber = vLockStatus(4)
            Call RemoteUnlockCRFPageInstance(TrialIdFromName(sClinicalTrialName), sTrialSite, lPersonId, lCRFPageTaskId, lVisitId, nVisitCycleNumber)
            'stop timing
            Call mDataTransferTime.IncrementLockFreeze
            Call mDataTransferTime.StopTiming(LockFreezeSection)
        Case ExchangeMessageType.LabDefinitionServerToSite
            On Error GoTo Timeout
            ' start timing
            Call mDataTransferTime.StartTiming(LaboratoryDownSection)
            ' set importok
            bImportOK = True
            
            sLaboratory = Mid(msMessageParameters, 1, InStr(msMessageParameters, "_") - 1)
            ' DPH Update display
            Call SetTextBoxDisplay("Downloading " & sLaboratory & " laboratory definition", trTransferData)
            sCabImportFile = msMessageParameters
            'call GetByteResponseHTTP which places contents of message into mbSDD
            ' DPH 02/05/2002 - Changed function so error handled gracefully
            'mbSDD() = GetByteResponseHTTP(Me.HTTPAddress & sCabImportFile, 300)
            mbSDD() = GetDataHTTP(Me.HTTPAddress & sCabImportFile, 300)
            mnFreeFile = FreeFile
            Open gsIN_FOLDER_LOCATION & sCabImportFile For Binary Access Write As #mnFreeFile
            Put #mnFreeFile, , mbSDD()
            Close #mnFreeFile
                
            On Error GoTo ErrHandler
                
            Set oExchange = New clsExchange
            'Unpack the CAB file into an LDD file,
            oExchange.ImportLDDCAB gsIN_FOLDER_LOCATION & sCabImportFile
            'get the name of the ldd file using the DIR command
            msLDDFile = Dir(gsCAB_EXTRACT_LOCATION & "*.ldd")
            If msLDDFile > "" Then
                If oExchange.ImportLDD(gsCAB_EXTRACT_LOCATION & msLDDFile) = ExchangeError.Success Then
                    'check to seee if there is an entry on Sitelaboratory for this lab
                    sSQL = "Select LaboratoryCode FROM SiteLaboratory WHERE LaboratoryCode ='" & sLaboratory & "'"
                    Set rsLabSites = New ADODB.Recordset
                    rsLabSites.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
                    'if no labsite entry exists then insert one
                    If rsLabSites.EOF Then
                        sSQL = "INSERT INTO SiteLaboratory (Site,LaboratoryCode) " _
                            & " VALUES ('" & Me.Site & "','" & sLaboratory & "')"
                        MacroADODBConnection.Execute sSQL
                    End If
                    rsLabSites.Close
                    Set rsLabSites = Nothing
                Else
                    ' ImportLDD error
                    bImportOK = False
                End If
                
                Set oExchange = Nothing
            Else
                ' No LDD file
                bImportOK = False
            End If
            
            'Kill the Cab File
            Kill gsIN_FOLDER_LOCATION & sCabImportFile

            ' stop timing
            Call mDataTransferTime.IncrementLaboratoryDown
            Call mDataTransferTime.StopTiming(LaboratoryDownSection)
            
            If bImportOK Then
                Call SetTextBoxDisplay(sLaboratory & " laboratory definition downloaded", trTransferData)
            Else
                Call SetTextBoxDisplay(sLaboratory & " laboratory definition failed to download. Please contact your systems administrator", trTransferData)
                ' Log Error
                gLog gsDOWNLOAD_MESG, sLaboratory & " laboratory definition failed to download. Please contact your systems administrator"
                mbTransferOK = False
            End If
            
        Case ""
            GoTo Timeout
        End Select

        ' Increment progress bar
        If (prgProgress.Value + nPercentInc) < 10 Then
            Call SetProgressBar(prgProgress.Value + nPercentInc, "")
        End If
        
        On Error GoTo Timeout
        
        'Note that msGetNextmessageURL is passed the messageId that has just been
        'successfully processed, for the purpose of setting it to read
        ' DPH 01/05/2002 - Changed call so throws error to be caught gracefully rather than MACRO error
        'msNextMessage = GetStringResponseHTTP(Me.HTTPAddress & msGetNextMessageURL & "?site=" & Me.Site & "&previousmessageid=" & msMessageId & "&trialoffice=" & Me.MACROServerDesc, mnDataTransferTimeout)
        msNextMessage = PostDataHTTP(Me.HTTPAddress & msGetNextMessageURL & "?site=" & Me.Site & "&previousmessageid=" & msMessageId & "&trialoffice=" & Me.MACROServerDesc, mnDataTransferTimeout)
    
        On Error GoTo ErrHandler
    
        ' DPH 01/05/2002 - Pick up if an error has occurred in the ASP page
        If Left(msNextMessage, 5) = "ERROR" Or InStr(1, LCase(msNextMessage), "<html") > 0 Then
            Call SetTextBoxDisplay("Messages could not be downloaded", trStepCompleted, False)
            ' log error
            gLog gsDOWNLOAD_MESG, "Messages could not be downloaded."
            mbTransferOK = False
            Exit Function
        End If
    
    Loop
    
    If Me.Visible = True Then
        
        cmdCancel.Caption = "&Cancel"
        cmdCancel.Visible = True

        Call SetTextBoxDisplay("All study messages from " & Me.MACROServerDesc & " downloaded", trStepCompleted)

    End If
    
    Me.MousePointer = vbDefault
    DownloadMessages = True
    
Exit Function

Timeout:
    'DisplayTimeoutMessage
    DownloadMessages = False
Exit Function

ErrHandler:
    Call SetConnectionStatus(csError)
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "DownLoadMessages", "frmDataTransfer")
          Case OnErrorAction.Ignore
              Resume Next
          Case OnErrorAction.Retry
              Resume
          Case OnErrorAction.QuitMACRO
              Call ExitMACRO
              End
     End Select
End Function


'---------------------------------------------------------------------
Private Function DownLoadMIMessages() As Boolean
'---------------------------------------------------------------------
'Mo Morris 22/11/00, Changed for new field MIMessageResponseTimeStamp
' NCJ 1 Mar 01 - Changed ConvertStandardToLocalNum to ConvertLocalNumToStandard in SQL
' NCJ 24 Oct 01 - Patched in changes made by MLM in 2.1:
'   Corrected regional settings problems.
'   Timestamps in double format are read from the server
'   and saved to the db in "standard" format, but stored "locally" in between.
' DPH 01/05/2002 - Replaced GetStringResponseHTTP with PostDataHTTP calls
'               Improved error handling on handling response data
' DPH 20/11/2002 - Stop Key Violations on previously received/not confirmed mimessages
' NCJ 20 Jan 03 - Update subject statuses on receipt of each message
'               NB This does NOT lock the subjects while it's doing it!!! (TO DO)
'---------------------------------------------------------------------
Dim sNextMessage As String
Dim lMIMessageID As Long
Dim sMIMessageSite As String
Dim nMIMessageSource As Integer
Dim nMIMessageType As Integer
Dim nMIMessageScope As Integer
Dim lMIMessageObjectID As Long
Dim nMIMessageObjectSource As Integer
Dim nMIMessagePriority As Integer
Dim sMIMessageTrialName As String
Dim lMIMessagePersonID As Long
Dim lMIMessageVisitId As Long
Dim nMIMessageVisitCycle As Integer
Dim lMIMessageCRFPageTaskID As Long
Dim lMIMessageResponseTaskID As Long
Dim sMIMessageResponseValue As String
Dim lMIMessageOCDiscrepancyID As Long
Dim dMIMessageCreated As Double
Dim dMIMessageSent As Double
Dim dMIMessageReceived As Double
Dim nMIMessageHistory As Integer
Dim nMIMessageProcessed As Integer
Dim nMIMessageStatus As Integer
Dim sMIMessageText As String
Dim sMIMessageUserName As String
Dim sMIMessageUserNameFull As String
Dim dMIMessageResponseTimeStamp As Double
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim bNotaPriorityChangeMessage As Boolean
' Used to control percentage increment
Dim nPercentInc As Integer
Dim bInsertedBefore As Boolean
' NCJ 15 Jan 03
Dim nMIMessageResponseCycle As Integer
Dim nMIMessageCreated_TZ As Integer
Dim nMIMessageReceived_TZ As Integer
Dim nMIMessageSent_TZ As Integer
Dim lMIMessageCRFPageCycle As Long
Dim lMIMessageCRFPageId As Long
Dim lMIMessageDataItemId As Long

' NCJ 20 Jan 03
Dim lStudyId As Long
Dim enMsgType As MIMsgType
Dim enMsgScope As MIMsgScope

    On Error GoTo ErrHandler
    
    Me.cmdCancel.Visible = True
    Call SetTextBoxDisplay("Downloading User messages from " & Me.MACROServerDesc, trTransferData)
    nPercentInc = 1
    
    On Error GoTo Timeout
    
    ' DPH 01/05/2002 - Changed call so throws error to be caught gracefully rather than MACRO error
    'sNextMessage = GetStringResponseHTTP(Me.HTTPAddress & msGetNextMIMessageURL & "?site=" & Me.Site, mnDataTransferTimeout)
    sNextMessage = PostDataHTTP(Me.HTTPAddress & msGetNextMIMessageURL & "?site=" & Me.Site, mnDataTransferTimeout)
    
    On Error GoTo ErrHandler

    ' DPH 01/05/2002 - Pick up if an error has occurred in the ASP page
    If Left(sNextMessage, 5) = "ERROR" Or InStr(1, LCase(sNextMessage), "<html") > 0 _
        Or (Len(sNextMessage) > 1 And InStr(1, sNextMessage, "<br>") = 0) Then
        Call SetTextBoxDisplay("User messages could not be downloaded", trStepCompleted, False)
        ' log error
        gLog gsDOWNLOAD_MIMESG, "User messages could not be downloaded."
        mbTransferOK = False
        Exit Function
    End If
    
    If sNextMessage <> "." Then
        ' start timing
        Call mDataTransferTime.StartTiming(MIMessagesDownSection)
    End If
    
    Do While sNextMessage <> "."
        'extract the individual fields from the user message
        lMIMessageID = CLng(ExtractFirstItemFromList(sNextMessage, "<br>"))
        sMIMessageSite = ExtractFirstItemFromList(sNextMessage, "<br>")
        nMIMessageSource = CInt(ExtractFirstItemFromList(sNextMessage, "<br>"))
        nMIMessageType = CInt(ExtractFirstItemFromList(sNextMessage, "<br>"))
        nMIMessageScope = CInt(ExtractFirstItemFromList(sNextMessage, "<br>"))
        lMIMessageObjectID = CLng(ExtractFirstItemFromList(sNextMessage, "<br>"))
        nMIMessageObjectSource = CInt(ExtractFirstItemFromList(sNextMessage, "<br>"))
        nMIMessagePriority = CInt(ExtractFirstItemFromList(sNextMessage, "<br>"))
        sMIMessageTrialName = ExtractFirstItemFromList(sNextMessage, "<br>")
        lMIMessagePersonID = CLng(ExtractFirstItemFromList(sNextMessage, "<br>"))
        lMIMessageVisitId = CLng(ExtractFirstItemFromList(sNextMessage, "<br>"))
        nMIMessageVisitCycle = CInt(ExtractFirstItemFromList(sNextMessage, "<br>"))
        lMIMessageCRFPageTaskID = CLng(ExtractFirstItemFromList(sNextMessage, "<br>"))
        lMIMessageResponseTaskID = CLng(ExtractFirstItemFromList(sNextMessage, "<br>"))
        sMIMessageResponseValue = ExtractFirstItemFromList(sNextMessage, "<br>")
        lMIMessageOCDiscrepancyID = CLng(ExtractFirstItemFromList(sNextMessage, "<br>"))
        ' NCJ 24/10/01 Convert standard to local
        dMIMessageCreated = CDbl(ConvertStandardToLocalNum(ExtractFirstItemFromList(sNextMessage, "<br>")))
        dMIMessageSent = CDbl(ConvertStandardToLocalNum(ExtractFirstItemFromList(sNextMessage, "<br>")))
        dMIMessageReceived = CDbl(ConvertStandardToLocalNum(ExtractFirstItemFromList(sNextMessage, "<br>")))
        
        nMIMessageHistory = CInt(ExtractFirstItemFromList(sNextMessage, "<br>"))
        nMIMessageProcessed = CInt(ExtractFirstItemFromList(sNextMessage, "<br>"))
        nMIMessageStatus = CInt(ExtractFirstItemFromList(sNextMessage, "<br>"))
        sMIMessageText = ExtractFirstItemFromList(sNextMessage, "<br>")
        sMIMessageUserName = ExtractFirstItemFromList(sNextMessage, "<br>")
        sMIMessageUserNameFull = ExtractFirstItemFromList(sNextMessage, "<br>")
        ' NCJ 24/10/01 Convert standard to local
        dMIMessageResponseTimeStamp = CDbl(ConvertStandardToLocalNum(ExtractFirstItemFromList(sNextMessage, "<br>")))
        
        ' NCJ 15 Jan 03 - New fields added for MACRO 3.0
        nMIMessageResponseCycle = CInt(ExtractFirstItemFromList(sNextMessage, "<br>"))
        nMIMessageCreated_TZ = CInt(ExtractFirstItemFromList(sNextMessage, "<br>"))
        nMIMessageReceived_TZ = CInt(ExtractFirstItemFromList(sNextMessage, "<br>"))
        nMIMessageSent_TZ = CInt(ExtractFirstItemFromList(sNextMessage, "<br>"))
        lMIMessageCRFPageCycle = CLng(ExtractFirstItemFromList(sNextMessage, "<br>"))
        lMIMessageCRFPageId = CLng(ExtractFirstItemFromList(sNextMessage, "<br>"))
        lMIMessageDataItemId = CLng(ExtractFirstItemFromList(sNextMessage, "<br>"))
        
        DoEvents
        
        'set the message processed flag to 0
        nMIMessageProcessed = 0
        
        'time stamp the message received flag
        dMIMessageReceived = IMedNow
        
        'initialise the Priority Change Message flag
        bNotaPriorityChangeMessage = True
        
        'Priority Change Messages.
        'Data Monitors on the Server are allowed to change the Priority of a Discrepancy, this
        'is done in conjunction with setting its message Sent field back to 0 on the Server,
        'which has the effect of causing the message to be retransmitted.
        'Check for a Discrepancy message with status Raised,
        'because it might be a priority change message that only requires
        'MIMessagePriority and MIMessageReceived to be updated
        If nMIMessageType = MIMsgType.mimtDiscrepancy And nMIMessageStatus = eDiscrepancyMIMStatus.dsRaised Then
            sSQL = "SELECT MIMessagePriority, MIMessageReceived FROM MIMessage" _
                & " WHERE MIMessageId = " & lMIMessageID _
                & " AND MIMessageSite ='" & sMIMessageSite & "'" _
                & " AND MIMessageSource = " & nMIMessageSource
            Set rsTemp = New ADODB.Recordset
            rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
            If rsTemp.RecordCount = 1 Then
                'Update the priority and set the new Received time
                rsTemp!MIMessagePriority = nMIMessagePriority
                ' NCJ 1 Mar 01 - Don't need conversion when assigning directly
'                rsTemp!MIMessageReceived = ConvertLocalnumtoStandard(CStr(dMIMessageReceived))
                rsTemp!MIMessageReceived = dMIMessageReceived
                rsTemp.Update
                'No more processing required for this priority changing message
                bNotaPriorityChangeMessage = False
            End If
            rsTemp.Close
            Set rsTemp = Nothing
        End If
            
        If bNotaPriorityChangeMessage Then
            ' DPH 20/11/2002 Check if this exact message has been received previously.
            '   If we are here it is not a PriorityChange Message
            '   If it has do not write to database again as will cause DB Key violation
            sSQL = "SELECT Count(*) AS MICount FROM MIMessage" _
                & " WHERE MIMessageId = " & lMIMessageID _
                & " AND MIMessageSite ='" & sMIMessageSite & "'" _
                & " AND MIMessageSource = " & nMIMessageSource
            Set rsTemp = New ADODB.Recordset
            rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
            If rsTemp(0).Value > 0 Then
                bInsertedBefore = True
            Else
                bInsertedBefore = False
            End If
            rsTemp.Close
            Set rsTemp = Nothing

            If Not bInsertedBefore Then
                'Set the history flag on the previous message of a Discrepancies or SDV.
                'Discrepancies and SDVs have a MessageObjectID, Messages and Notes do not
                If lMIMessageObjectID > 0 Then
                    'Filtering on site not necessary on a Client/TriallOffice installation
                    sSQL = "UPDATE MIMessage " _
                        & " SET MIMessageHistory = " & MIMsgHistory.mimhNotCurrent _
                        & " WHERE MIMessageObjectId = " & lMIMessageObjectID _
                        & " AND MIMessageObjectSource = " & nMIMessageObjectSource _
                        & " AND MIMessageHistory = " & MIMsgHistory.mimhCurrent
                    MacroADODBConnection.Execute sSQL
                End If
                
                ' Store the new message
                ' NCJ 24/10/01 Convert dMIMessageResponseTimeStamp to standard
                ' NCJ 15 Jan 03 - New fields added for MACRO 3.0
                sSQL = "INSERT INTO MIMessage (MIMessageID,MIMessageSite,MIMessageSource,MIMessageType," _
                    & "MIMessageScope,MIMessageObjectID,MIMessageObjectSource,MIMessagePriority," _
                    & "MIMessageTrialName,MIMessagePersonID,MIMessageVisitId,MIMessageVisitCycle," _
                    & "MIMessageCRFPageTaskID,MIMessageResponseTaskID,MIMessageResponseValue," _
                    & "MIMessageOCDiscrepancyID,MIMessageCreated,MIMessageSent,MIMessageReceived," _
                    & "MIMessageHistory,MIMessageProcessed,MIMessageStatus,MIMessageText," _
                    & "MIMessageUserName,MIMessageUserNameFull,MIMessageResponseTimeStamp, " _
                    & "MIMessageResponseCycle,MIMessageCreated_TZ,MIMessageReceived_TZ, " _
                    & "MIMessageSent_TZ,MIMessageCRFPageCycle,MIMessageCRFPageId, MIMessageDataItemId)"
               sSQL = sSQL & "VALUES (" & lMIMessageID & ",'" & sMIMessageSite & "'," & nMIMessageSource & "," & nMIMessageType & "," _
                    & nMIMessageScope & "," & lMIMessageObjectID & "," & nMIMessageObjectSource & "," & nMIMessagePriority & ",'" _
                    & sMIMessageTrialName & "'," & lMIMessagePersonID & "," & lMIMessageVisitId & "," & nMIMessageVisitCycle & "," _
                    & lMIMessageCRFPageTaskID & "," & lMIMessageResponseTaskID & ",'" & ReplaceQuotes(sMIMessageResponseValue) & "'," _
                    & lMIMessageOCDiscrepancyID & "," & ConvertLocalNumToStandard(CStr(dMIMessageCreated)) & "," _
                    & ConvertLocalNumToStandard(CStr(dMIMessageSent)) & "," & ConvertLocalNumToStandard(CStr(dMIMessageReceived)) & "," _
                    & nMIMessageHistory & "," & nMIMessageProcessed & "," & nMIMessageStatus & ",'" _
                    & ReplaceQuotes(sMIMessageText) & "','" & sMIMessageUserName & "','" & ReplaceQuotes(sMIMessageUserNameFull) & "'," _
                    & ConvertLocalNumToStandard(CStr(dMIMessageResponseTimeStamp)) & ","
                sSQL = sSQL & nMIMessageResponseCycle & "," & nMIMessageCreated_TZ & "," & nMIMessageReceived_TZ & "," _
                    & nMIMessageSent_TZ & "," & lMIMessageCRFPageCycle & "," & lMIMessageCRFPageId & "," & lMIMessageDataItemId
                sSQL = sSQL & ")"
                
                MacroADODBConnection.Execute sSQL
                
                ' NCJ 20 Jan 03 - Now update the MIMStatuses for the subject
                ' TO DO - Sort out what we do if someone's got the subject open....
                enMsgType = nMIMessageType
                enMsgScope = nMIMessageScope
                lStudyId = TrialIdFromName(sMIMessageTrialName)
                Select Case enMsgType
                Case MIMsgType.mimtDiscrepancy, MIMsgType.mimtSDVMark
                    ' Some values will be 0 if scope is not Question
                    Call UpdateMIMsgStatus(goUser.CurrentDBConString, enMsgType, _
                                        sMIMessageTrialName, lStudyId, sMIMessageSite, lMIMessagePersonID, _
                                        lMIMessageVisitId, nMIMessageVisitCycle, _
                                        lMIMessageCRFPageTaskID, lMIMessageResponseTaskID, nMIMessageResponseCycle)
                Case MIMsgType.mimtNote
                    Call UpdateNoteStatus(goUser.CurrentDBConString, enMsgScope, _
                                        sMIMessageTrialName, lStudyId, sMIMessageSite, lMIMessagePersonID, _
                                        lMIMessageVisitId, nMIMessageVisitCycle, _
                                        lMIMessageCRFPageTaskID, lMIMessageResponseTaskID, nMIMessageResponseCycle)
                End Select
            End If
        End If
            
        On Error GoTo Timeout
                
        'Note that msGetNextMIMessageURL is passed the MIMessageID that has just been
        'successfully processed, for the purpose of setting it to sent (together with its sent time)
        ' NCJ 24/10/01 - Convert dMIMessageSent to standard
'        sNextMessage = GetStringResponse(Me.HTTPAddress & msGetNextMIMessageURL & "?site=" & Me.Site & "&PreviousMessageID=" & lMIMessageID & "&PreviousMessageSent=" & dMIMessageSent, mnDataTransferTimeout)
        ' DPH 01/05/2002 - Changed call so throws error to be caught gracefully rather than MACRO error
        sNextMessage = PostDataHTTP(Me.HTTPAddress & msGetNextMIMessageURL _
                        & "?site=" & Me.Site _
                        & "&PreviousMessageID=" & lMIMessageID _
                        & "&PreviousMessageSent=" & ConvertLocalNumToStandard(CStr(dMIMessageSent)), mnDataTransferTimeout)
        
        On Error GoTo ErrHandler
        
        ' DPH 01/05/2002 - Improve error handling
        If Left(sNextMessage, 5) = "ERROR" Or InStr(1, LCase(sNextMessage), "<html") > 0 _
        Or (Len(sNextMessage) > 1 And InStr(1, sNextMessage, "<br>") = 0) Then
            Call SetTextBoxDisplay("Some user messages could not be downloaded", trStepCompleted, False)
            ' log error
            gLog gsDOWNLOAD_MIMESG, "Some user messages could not be downloaded."
            mbTransferOK = False
            Exit Function
        End If

        ' Increment progress bar / mimessage timer
        Call mDataTransferTime.IncrementMIMessagesDown
        If (prgProgress.Value + nPercentInc) < 20 Then
            Call SetProgressBar(prgProgress.Value + nPercentInc, "")
        End If
    Loop
    
    If Me.Visible = True Then
        Me.cmdCancel.Caption = "&Cancel"
        Me.cmdCancel.Visible = True
        Call SetTextBoxDisplay("All user messages from " & Me.MACROServerDesc & " downloaded", trStepCompleted)
    End If
        
    Me.MousePointer = vbDefault
    
    DownLoadMIMessages = True
    
Exit Function
Timeout:
    'DisplayTimeoutMessage
    DownLoadMIMessages = False
    
Exit Function
ErrHandler:
  'REM 07/04/03 - Changed error handling so error is recordered and data transfer continues
  mbTransferOK = False
  Call SetTextBoxDisplay("An error was encountered during Download MIMessages" & vbCrLf & "Error Description: " & Err.Description & "Error Number: " & Err.Number, trStepCompleted, mbTransferOK)
  DownLoadMIMessages = True 'set to true so data transfer can continue as error was only local
End Function

'---------------------------------------------------------------------
Private Function SendMIMessages() As Boolean
'---------------------------------------------------------------------
'Mo Morris 22/11/00, Changed for new field MIMessageResponseTimeStamp
' NCJ 24/10/01 - Added in MLM regional settings changes from 2.1
' DPH 17/1/2002 - Replaced UserCode with UserNameFull
' DPH 25/03/2002 - Remove Me.Show as now modal
' DPH 10/06/2002 - Added information to Sending Notes/SDVs/Discrepancies CBBL 2.2.14.4
' REM 01/12/03 - Place RemoveNull around MIMessageText field
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim i As Integer
Dim sResultOfPosting As String
Dim dblSentTime As Double
' Used to control percentage increment
Dim nPercentInc As Integer
Dim sHTTPAddress As String
Dim sHTTPData As String

Dim sMessageInfo As String


Dim oTimezone As TimeZone

    On Error GoTo ErrHandler
    
    Me.cmdCancel.Visible = True

    Set oTimezone = New TimeZone
    
    sSQL = "SELECT * FROM MIMessage " _
    & "WHERE MIMessageSource = " & TypeOfInstallation.RemoteSite _
    & " AND MIMessageSent = 0"

    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        Call SetTextBoxDisplay("No user messages to be sent to " & Me.MACROServerDesc, trTransferData)
    Else
    
        Call SetTextBoxDisplay("Sending user messages to " & Me.MACROServerDesc, trTransferData)
        
        ' start timing
        Call mDataTransferTime.StartTiming(MIMessagesUpSection)
        
        For i = 1 To rsTemp.RecordCount
        
            ' calculate percentage to increase each mimessage
            ' sending mimessage 15%
            nPercentInc = CInt(15 \ rsTemp.RecordCount)
            If nPercentInc = 0 Then
                nPercentInc = 1
            End If
            
            On Error GoTo Timeout
            
            'store the time sent for the purpose of using the same time when updating it in the
            'sending Client database
            dblSentTime = IMedNow
            
            'changed Mo Morris 26/2/01, URLCharToHexEncoding now called for MIMessageResponseValue
'            SetTransferPropertiesHTTP Me.HTTPAddress & msReceiveMIMessageURL
            ' RS 30/09/2002:    Added Timezone values
            sHTTPAddress = Me.HTTPAddress & msReceiveMIMessageURL
            ' DPH 17/1/2002 - Replaced UserCode with UserNameFull
            ' REM 01/12/03 - Place RemoveNull around MIMessageText field
            sHTTPData = "ID=" & rsTemp!MIMessageID _
                & "&Site=" & rsTemp!MIMessageSite & "&Source=" & rsTemp!MIMessageSource _
                & "&Type=" & rsTemp!MIMessageType & "&Scope=" & rsTemp!MIMessageScope _
                & "&ObjectID=" & rsTemp!MIMessageObjectId & "&ObjectSource=" & rsTemp!MIMessageObjectSource _
                & "&Priority=" & rsTemp!MIMessagePriority & "&TrialName=" & rsTemp!MIMessageTrialName _
                & "&PersonID=" & rsTemp!MIMessagePersonId & "&VisitID=" & rsTemp!MIMessageVisitId _
                & "&VisitCycle=" & rsTemp!MIMessageVisitCycle & "&CRFPageTaskID=" & rsTemp!MIMessageCRFPageTaskId _
                & "&ResponseTaskID=" & rsTemp!MIMessageResponseTaskId & "&ResponseValue=" & URLCharToHexEncoding(rsTemp!MIMessageResponseValue) _
                & "&OCDiscrepancyID=" & rsTemp!MIMessageOCDiscrepancyID & "&Created=" & ConvertLocalNumToStandard(CStr(rsTemp!MIMessageCreated)) _
                & "&Sent=" & ConvertLocalNumToStandard(CStr(dblSentTime)) & "&Received=" & ConvertLocalNumToStandard(CStr(rsTemp!MIMessageReceived)) _
                & "&History=" & rsTemp!MIMessageHistory & "&Processed=" & rsTemp!MIMessageProcessed _
                & "&Status=" & rsTemp!MIMessageStatus & "&Text=" & URLCharToHexEncoding(RemoveNull(rsTemp!MIMessageText)) _
                & "&UserName=" & URLCharToHexEncoding(rsTemp!MIMessageUserName) & "&UserNameFull=" & URLCharToHexEncoding(rsTemp!MIMessageUserNameFull) _
                & "&ResponseTimeStamp=" & ConvertLocalNumToStandard(CStr(rsTemp!MIMessageResponseTimeStamp))
            ' NCJ 15 Jan 03 - New fields added for MACRO 3.0
            sHTTPData = sHTTPData & "&Created_TZ=" & rsTemp!MIMessageCreated_TZ _
                & "&Sent_TZ=" & oTimezone.TimezoneOffset _
                & "&Received_TZ=" & rsTemp!MIMessageReceived_TZ _
                & "&ResponseCycle=" & rsTemp!MIMessageResponseCycle _
                & "&CRFPageCycle=" & rsTemp!MIMessageCRFPageCycle _
                & "&CRFPageID=" & rsTemp!MIMessageCRFPageId _
                & "&DataItemID=" & rsTemp!MIMessageDataItemId
            
            ' New HTTP control
            
'            Debug.Print sHTTPAddress & ": " & sHTTPData
            sResultOfPosting = PostDataHTTP(sHTTPAddress, 90, sHTTPData)
            
            DoEvents
            
            On Error GoTo ErrHandler
            
            If sResultOfPosting = "SUCCESS" Then
                'set MIMessageSent = now for the message that has just been successfuly sent to the server
                rsTemp!MIMessageSent = dblSentTime
                rsTemp!MIMessageSent_TZ = oTimezone.TimezoneOffset
                rsTemp.Update
                ' DPH 10/06/2002 - Added information to Sending Notes/SDVs/Discrepancies CBBL 2.2.14.4
                sMessageInfo = GetMIMTypeText(rsTemp!MIMessageType)
                If sMessageInfo <> "" Then
                    sMessageInfo = sMessageInfo & " for Subject " & rsTemp!MIMessageSite & "/" & rsTemp!MIMessagePersonId
                    sMessageInfo = sMessageInfo & " sent to " & Me.MACROServerDesc & " successfully"
                Else
                    sMessageInfo = "Invalid Message type sent to server. Please contact your MACRO systems administrator."
                End If
                Call SetTextBoxDisplay(sMessageInfo, trTransferData)
            Else
                Call SetTextBoxDisplay("User messages could not be sent to " & Me.MACROServerDesc, trTransferData)
                mbTransferOK = False
                'the last message was not read successfully so exit the loop by moving past the last record
                rsTemp.MoveLast
            End If
            ' update timing class
            Call mDataTransferTime.IncrementMIMessagesUp
            ' Increment progress bar
            If (prgProgress.Value + nPercentInc) < 66 Then
                Call SetProgressBar(prgProgress.Value + nPercentInc, "")
            End If
            rsTemp.MoveNext
        Next
        Call SetTextBoxDisplay("User messages have been sent to " & Me.MACROServerDesc, trStepCompleted)
        
    End If
    rsTemp.Close
    Set rsTemp = Nothing
    SendMIMessages = True
Exit Function

Timeout:
    DisplayTimeoutMessage
    SendMIMessages = False
Exit Function

ErrHandler:
  'REM 07/04/03 - Changed error handling so error is recordered and data transfer continues
  mbTransferOK = False
  Call SetTextBoxDisplay("An error was encountered during Send MIMessages" & vbCrLf & "Error Description: " & Err.Description & "Error Number: " & Err.Number, trStepCompleted, mbTransferOK)
  SendMIMessages = True
'  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "SendMIMessages", "frmDataTransfer")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            End
'   End Select
End Function

'---------------------------------------------------------------------
Private Function SendLaboratory() As Boolean
'---------------------------------------------------------------------
' REVISIONS
' DPH 15/04/2002 - CAB Validation & Checksum added
'REM 08/04/03 - changed to a function
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim i As Integer
Dim sResultOfPosting As String
Dim sCabFileName As String
Dim nFileNumber As Integer
Dim sFullFileName As String
Dim sFileContent As String
Dim sHTTPAddress As String
Dim sHTTPData As String

' DPH 15/04/2002 - Add collection of files / filesizes for CAB validation
Dim colCABContents As New Collection
Dim sCABChecksum As String
Dim oMACROCheckSum As CheckSum
Dim dTime As Double
Dim nTry As Integer

    On Error GoTo ErrHandler
    
    Set oMACROCheckSum = New CheckSum
    
    Me.cmdCancel.Visible = True
    
    Set oExchange = New clsExchange
    
    sSQL = "SELECT * FROM Laboratory WHERE Changed = " & Changed.Changed
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        Call SetTextBoxDisplay("No Laboratory Definition files to be sent to " & Me.MACROServerDesc, trTransferData)
    Else
        Call SetTextBoxDisplay("Sending Laboratory data to " & Me.MACROServerDesc, trTransferData)
        ' start timing
        Call mDataTransferTime.StartTiming(LaboratoryUpSection)
        
        Call SetTextBoxDisplay("Laboratory Definition files being sent to " & Me.MACROServerDesc, trTransferData)
        For i = 1 To rsTemp.RecordCount
        
            sCabFileName = oExchange.ExportLDD(rsTemp!LaboratoryCode, colCABContents)
            
            If sCabFileName > "" And GetFileLength(gsOUT_FOLDER_LOCATION & sCabFileName) > 0 Then
                
                'convert cab file to a HEX coded file that can be transfered using HTTP
                HEXEncodeFileXCeed gsOUT_FOLDER_LOCATION & sCabFileName, gsOUT_FOLDER_LOCATION & "_" & sCabFileName
'                HEXEncodeFile gsOUT_FOLDER_LOCATION & sCabFileName, gsOUT_FOLDER_LOCATION & "_" & sCabFileName
                
                ' DPH 25/03/2002 - Validate CAB file
                If Not ValidateZIP(gsOUT_FOLDER_LOCATION & "_" & sCabFileName, colCABContents) Then
                    ' Log Error
                    gLog gsVALIDATE_ZIP, "CAB " & "_" & sCabFileName & " failed to validate correctly."
                    ' message on screen
                    Call SetTextBoxDisplay("Data not sent in Laboratory Definition file " & "_" & sCabFileName & " as validation of data file failed", trTransferData)
                    mbTransferOK = False
                Else
                    ' Checksum Hex CAB & send
                    sCABChecksum = oMACROCheckSum.GetFileCheckSum(gsOUT_FOLDER_LOCATION & "_" & sCabFileName)
                
'                    Call SetTimeoutHTTP(300)

                    nFileNumber = FreeFile
                    sFullFileName = gsOUT_FOLDER_LOCATION & "_" & sCabFileName
                    Open sFullFileName For Input As nFileNumber
                    sFileContent = Input(FileLen(sFullFileName), #nFileNumber)
                    
                    On Error GoTo Timeout
                    
                    ' DPH 15/04/2002 - Added Checksum to ASP call
                    sHTTPAddress = Me.HTTPAddress & msReceiveLaboratoryURL
                    sHTTPData = "LaboratoryCode=" & rsTemp!LaboratoryCode _
                        & "&Site=" & Me.Site & "&FileName=" & sCabFileName & "&Data=" & _
                        sFileContent & "&chksum=" & sCABChecksum
                    
'                    Me.Inet1.Execute , "POST", "LaboratoryCode=" & rsTemp!LaboratoryCode _
'                        & "&Site=" & Me.Site & "&FileName=" & sCabFileName & "&Data=" & _
'                        sFileContent & "&chksum=" & sCABChecksum, "Content-Type: application/x-www-form-urlencoded"
                    Close nFileNumber
                                        
                    sResultOfPosting = PostDataHTTP(sHTTPAddress, 300, sHTTPData)
                    
                    DoEvents
                    
                    On Error GoTo ErrHandler
                    
                    Select Case sResultOfPosting
                    
                        Case msASP_RETURN_SUCCESS
                        
                            ' write message to screen
                            Call SetTextBoxDisplay("Laboratory data sent successfully in data file " & "_" & sCabFileName & " to " & Me.MACROServerDesc, trTransferData)
                            
                            'Update the changed flag for the successfully transfered Laboratory
                            rsTemp!Changed = Changed.NoChange
                            
                        Case msASP_RETURN_CHECKSUM_ERROR
                        
                            ' write message to screen
                            Call SetTextBoxDisplay("Error in Checksum data file " & "_" & sCabFileName & " sending to " & Me.MACROServerDesc, trTransferData)
                            
                            ' Log error
                            gLog gsVALIDATE_ZIP, "CAB " & "_" & sCabFileName & " failed to validate correctly on " & Me.MACROServerDesc
                        
                            mbTransferOK = False
                            
                        Case Else
                        
                            Call SetTextBoxDisplay("Laboratory Definition file could not be sent to " & Me.MACROServerDesc, trTransferData)
                            'the laboratory defintion has not been read by the server so exit the loop by moving past last record
                            rsTemp.MoveLast
                    
                            mbTransferOK = False
                            
                    End Select

                End If
            Else
                ' Show Failed message
                Call SetTextBoxDisplay("Laboratory Definition file " & "_" & sCabFileName & " has not been sent as failed to create data file.", trTransferData)
                mbTransferOK = False
            End If
            ' update timer counter
            Call mDataTransferTime.IncrementLaboratoryUp
            
            rsTemp.MoveNext
        Next
        Call SetTextBoxDisplay("Laboratory Definition files have been sent to " & Me.MACROServerDesc, trStepCompleted)
    End If
    rsTemp.Close
    Set rsTemp = Nothing
    Set oExchange = Nothing
    Set oMACROCheckSum = Nothing
    
    SendLaboratory = True
    
Exit Function

Timeout:
    DisplayTimeoutMessage
    SendLaboratory = False
Exit Function

ErrHandler:
  mbTransferOK = False
  Call SetTextBoxDisplay("An error was encountered while sending laboratory data to the server" & vbCrLf & "Error Description: " & Err.Description & "Error Number: " & Err.Number, trStepCompleted, mbTransferOK)
  SendLaboratory = True 'set to true as can still continue with the rest of data transfer
'  Call SetConnectionStatus(csError)
'  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "SendLaboratory", "frmDataTransfer")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            Call MACROEnd
'   End Select
End Function

'---------------------------------------------------------------------
Public Function PerformIntegrityCheck() As Boolean
'---------------------------------------------------------------------
' Create and transmit data integrity checks to the server
'---------------------------------------------------------------------
' DPH 15/04/2002 - Added Data integrity routine
' REM 08/04/03 - Changed to a function
'---------------------------------------------------------------------

Dim sSQL As String
Dim rsSubject As ADODB.Recordset
Dim rsMIMessage As ADODB.Recordset
Dim sResultOfPosting As String
Dim dblSentTime As Double
Dim sTransmission As String
Dim dblTime As Double
Dim nTry As Integer
' Data collection variables
Dim dblSumTimeStamp As Double
Dim dblMaxTimeStamp As Double
Dim lCountSubjectRec As Long
Dim dblSumMessageIdZero As Double
Dim dblSumMessageIdOne As Double
Dim lCountMessageIdZero As Long
Dim lCountMessageIdOne As Long
Dim nSource As Integer
Dim sHTTPAddress As String
Dim sHTTPData As String

    On Error GoTo ErrHandler
    
    Me.cmdCancel.Visible = True

    Call SetTextBoxDisplay("Sending Data Integrity Information to " & Me.MACROServerDesc & "...", trTransferData)
    DoEvents

    ' Subject Data
    sSQL = "SELECT SUM(ResponseTimeStamp) AS SumResponse, Max(ResponseTimeStamp) AS MaxResponse, Count(ResponseTimeStamp) AS CountResponse " & _
    "FROM DataItemResponseHistory, CRFElement " & _
    "WHERE ((DataItemResponseHistory.ClinicalTrialId = CRFElement.ClinicalTrialId) AND (DataItemResponseHistory.CRFPageID = CRFElement.CRFPageID) " & _
    "AND (DataItemResponseHistory.CRFElementId = CRFElement.CRFElementId)) " & _
    "AND (CRFElement.LocalFlag = 0)"
       
    Set rsSubject = New ADODB.Recordset
    rsSubject.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    ' initialise variables
    dblSumTimeStamp = 0
    dblMaxTimeStamp = 0
    lCountSubjectRec = 0
    
    If Not rsSubject.EOF Then
        ' Get data from recordset
        ' Sum of ResponseTimeStamp
        If IsNull(rsSubject(0).Value) Then
            dblSumTimeStamp = 0
        Else
            dblSumTimeStamp = rsSubject(0).Value
        End If
        ' Max ResponseTimeStamp
        If IsNull(rsSubject(1).Value) Then
            dblMaxTimeStamp = 0
        Else
            dblMaxTimeStamp = rsSubject(1).Value
        End If
        ' Count Of Records
        If IsNull(rsSubject(2).Value) Then
            lCountSubjectRec = 0
        Else
            lCountSubjectRec = rsSubject(2).Value
        End If
    End If
    
    rsSubject.Close

    ' Create Transmission String
    sTransmission = "?DIVTYPE=History&SITENAME=" & Me.Site & "&DIVSUBJECT=" & CStr(dblSumTimeStamp) & "&MAXTIMESTAMP=" & CStr(dblMaxTimeStamp)
    sTransmission = sTransmission & "&SUBJECTCOUNT=" & lCountSubjectRec & "&USERID=" & goUser.UserName
    
    ' send to ASP
    On Error GoTo Timeout
        
    ' Subject Data Transmission
    sHTTPAddress = Me.HTTPAddress & msDataIntegrityURL
    sHTTPData = sTransmission
    sResultOfPosting = PostDataHTTP(sHTTPAddress & sHTTPData, 90)
        
    DoEvents
    
    On Error GoTo ErrHandler
    
    If sResultOfPosting = "SUCCESS" Then
        ' Successful receive by server
        Call SetTextBoxDisplay("Subject data integrity check sent to " & Me.MACROServerDesc, trTransferData)
    Else
        Call SetTextBoxDisplay(msSUBJECT_DATA_INTEGRITY_ERROR & Me.MACROServerDesc, trTransferData)
        ' Mark client log as a fail
        gLog gsDATA_INTEG_COMMS, msSUBJECT_DATA_INTEGRITY_ERROR & Me.MACROServerDesc
        mbTransferOK = False
    End If
        
    ' MIMessage Data
    sSQL = "SELECT SUM(MIMessageId) AS SumId, MIMessageSource, Count(MIMessageId) AS CountId " & _
            "FROM MIMessage GROUP BY MIMessageSource"

    Set rsMIMessage = New ADODB.Recordset
    rsMIMessage.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    ' initialise message ID sums
    dblSumMessageIdZero = 0
    dblSumMessageIdOne = 0
    lCountMessageIdZero = 0
    lCountMessageIdOne = 0
    
    Do While Not rsMIMessage.EOF
        Select Case rsMIMessage(1).Value
            Case 0
                dblSumMessageIdZero = rsMIMessage(0).Value
                lCountMessageIdZero = rsMIMessage(2).Value
            Case 1
                dblSumMessageIdOne = rsMIMessage(0).Value
                lCountMessageIdOne = rsMIMessage(2).Value
        End Select
        rsMIMessage.MoveNext
    Loop
    
    rsMIMessage.Close
        
    ' Transmit MIMessage info data to server
    ' Create Transmission String
    sTransmission = "?DIVTYPE=MIMessage&SITENAME=" & Me.Site & "&DIVIDZERO=" & CStr(dblSumMessageIdZero) & "&DIVIDONE=" & CStr(dblSumMessageIdOne)
    sTransmission = sTransmission & "&COUNTIDZERO=" & lCountMessageIdZero & "&COUNTIDONE=" & lCountMessageIdOne & "&USERID=" & goUser.UserName
    
    ' send to ASP
    On Error GoTo Timeout
            
    sHTTPAddress = Me.HTTPAddress & msDataIntegrityURL
    sHTTPData = sTransmission
    sResultOfPosting = PostDataHTTP(sHTTPAddress & sHTTPData, 240)
    
    DoEvents
    
    On Error GoTo ErrHandler
    
    If sResultOfPosting = "SUCCESS" Then
        ' Successful receive by server
        Call SetTextBoxDisplay("MIMessage Data integrity check sent to " & Me.MACROServerDesc, trTransferData)
    Else
        Call SetTextBoxDisplay(msMIMESSAGE_DATA_INTEGRITY_ERROR & Me.MACROServerDesc, trTransferData)
        ' Mark client log as a fail
        gLog gsDATA_INTEG_COMMS, msMIMESSAGE_DATA_INTEGRITY_ERROR & Me.MACROServerDesc
        mbTransferOK = False
    End If

    Call SetTextBoxDisplay("Data integrity information sent to " & Me.MACROServerDesc, trTransferData)
    
    Set rsSubject = Nothing
    Set rsMIMessage = Nothing
    
    PerformIntegrityCheck = True
    
Exit Function

Timeout:
    DisplayTimeoutMessage
    PerformIntegrityCheck = False
Exit Function

ErrHandler:
  mbTransferOK = False
  Call SetTextBoxDisplay("An error was encountered while sending the Integrity check to the server" & vbCrLf & "Error Description: " & Err.Description & "Error Number: " & Err.Number, trStepCompleted, mbTransferOK)
  PerformIntegrityCheck = True 'set to true as can still continue with the rest of data transfer
'  Call SetConnectionStatus(csError)
'  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "PerformIntegrityCheck", "frmDataTransfer")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Call ExitMACRO
'            End
'   End Select
End Function

'---------------------------------------------------------------------
Public Function ValidateZIP(sHexZIPFileName As String, ByVal colZIPContents As Collection) As Boolean
'---------------------------------------------------------------------
' Validate CAB file before sending it to the server
'---------------------------------------------------------------------
Dim bPassed As Boolean
Dim sZipFileName As String
Dim sNextExtractedFile As String
Dim lFileSize As Long
Dim i As Integer

On Error GoTo ErrHandler

    ' initialise variables
    bPassed = True
    
    ' Initialise CAB Filename
    sZipFileName = gsIN_FOLDER_LOCATION & "MACROCABVAL.zip"
    
    ' Firstly UnHex CAB
    If HEXDecodeFile(sHexZIPFileName, sZipFileName) Then
            
        ' Extract files to CABExtract Folder
        oExchange.ClearCabExtractFolder
        
        ' Make sure file exists before opening
        If FolderExistence(sZipFileName, True) Then
            
            ' Error handler for Unzip process
            On Error GoTo UnZipErr
            
            ' Extract Files from ZIP using XCeed
            Call UnZipFiles(gsCAB_EXTRACT_LOCATION, sZipFileName)
        
            On Error GoTo ErrHandler
            
            ' Check File list extracted with that in collection
            sNextExtractedFile = Dir(gsCAB_EXTRACT_LOCATION & "*.*")
            Do While sNextExtractedFile <> ""
                ' get file size
                lFileSize = GetFileLength(gsCAB_EXTRACT_LOCATION & sNextExtractedFile)
                
                ' StripFileNameFromPath
                If CollectionMember(colZIPContents, sNextExtractedFile, False) Then
                    ' compare file lengths
                    If colZIPContents(sNextExtractedFile) = lFileSize Then
                        ' This file passed so remove
                        colZIPContents.Remove (sNextExtractedFile)
                    Else
                        gLog gsVALIDATE_ZIP, "ZIP Validation Error. File " & sNextExtractedFile & " has a different length to the original file in CAB. "
                        bPassed = False
                    End If
                End If
            
                'get next prd file via the DIR command
                sNextExtractedFile = Dir
            Loop
            
            If colZIPContents.Count > 0 Then
                gLog gsVALIDATE_ZIP, "ZIP Validation Error. Not all files were correctly cabbed in CAB " & sHexZIPFileName
                bPassed = False
            End If
            
        Else
            gLog gsVALIDATE_ZIP, "ZIP Validation Error. Hex decode of " & sHexZIPFileName & " failed. "
            bPassed = False
        End If
        
    Else
        gLog gsVALIDATE_ZIP, "ZIP Validation Error. Hex decode of " & sHexZIPFileName & " failed. "
        bPassed = False
    End If
    
    ' Delete Temporary Files
    ' Make sure file exists before deleting
    If FolderExistence(sZipFileName, True) Then
        Kill sZipFileName
    End If
    oExchange.ClearCabExtractFolder
   
    ' Pass back function result
    ValidateZIP = bPassed
    
    Exit Function

UnZipErr:
    ' If unzipping fails
    gLog gsVALIDATE_ZIP, "Extracting " & StripFileNameFromPath(sHexZIPFileName) & " failed. XCeed Error Number " & Err.Number & ". "
    ValidateZIP = False
    
    Exit Function
ErrHandler:
  Call SetConnectionStatus(csError)
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ValidateZIP", "frmDataTransfer")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

''---------------------------------------------------------------------
'Private Sub SetTransferProperties(ByVal sAddress As String)
''---------------------------------------------------------------------
''   ATN 14/3/2000
''   Need to set transfer properties each time
'' REVISIONS
'' DPH 17/1/2002 - Changed order of Inet properties being set so port no is taken into account
''---------------------------------------------------------------------
'
'    ' DPH 17/1/2002 - Set URL Firstly rather than last so Port gets 'added' to URL rather than overwritten
'    Inet1.URL = sAddress
'
'    Inet1.UserName = frmMenu.gTrialOffice.User
'    Inet1.Password = frmMenu.gTrialOffice.Password
'    Inet1.RemotePort = frmMenu.gTrialOffice.PortNumber
'    Inet1.Proxy = frmMenu.gTrialOffice.ProxyServer
'
''    Inet1.URL = sAddress
'End Sub

'---------------------------------------------------------------------
Public Function AddPortToHTTPString(sHTTP As String, nPort As Integer) As String
' Tacks port onto http reference call
' Assumes port not already in http ref
'---------------------------------------------------------------------
' REVISIONS
'---------------------------------------------------------------------

Dim nPos As Integer
Dim sNewHTTP As String

    ' Default to original http
    sNewHTTP = sHTTP
    
    ' Get position of end of http:// reference
    nPos = InStr(1, sHTTP, "//", vbTextCompare)
    If nPos > 0 Then
        ' Move to character after //
        nPos = nPos + 2
        ' Get end of HostString
        nPos = InStr(nPos, sHTTP, "/", vbTextCompare)
        ' Tag port number at end of hoststring & before directories
        ' http://hostname:port/dir/
        If nPos > 0 And nPort > 0 Then
            sNewHTTP = Left(sHTTP, nPos - 1) & ":" & nPort & Right(sHTTP, Len(sHTTP) - nPos + 1)
        End If
    End If
    
    AddPortToHTTPString = sNewHTTP
    
End Function

''---------------------------------------------------------------------
'Private Function GetStringResponse(ByVal vAddress As String, _
'                                   ByVal vTimeout As Integer) As String
''---------------------------------------------------------------------
''Called by:-    DownLoadMessages
''               DisplayGetMessages
''               Register
''               Randomise
''Using Microsofts Internet Control using HTTP this sub reads from a
''TrialOffice server using POST and GetChunk
''---------------------------------------------------------------------
'
'Dim vtData As Variant ' Data variable.
'Dim strData As String
'Dim bDone As Boolean
'
'On Error GoTo ErrHandler
'
'    strData = ""
'    bDone = False
'
'    Inet1.RequestTimeout = vTimeout
'
'    '   ATN 14/2/2000
'    '   Need to set transfer properties each time the control is executed
'    SetTransferProperties vAddress
'    Inet1.Execute , "POST"
'    DoEvents
'    Inet1.Cancel
'    DoEvents
'    '   ATN 14/2/2000
'    SetTransferProperties vAddress
'    Inet1.Execute , "POST"
'    DoEvents
'    Do While Inet1.StillExecuting
'        DoEvents
'    Loop
'
'    vtData = Inet1.GetChunk(4096, icString)
'
'    DoEvents
'
'    Do While Not bDone
'
'        strData = strData & vtData
'        ' Get next chunk.
'        vtData = Inet1.GetChunk(4096, icString)
'        DoEvents
'        If Len(vtData) = 0 Then
'            bDone = True
'        End If
'    Loop
'
'    GetStringResponse = strData
'
'Exit Function
'ErrHandler:
'  Call SetConnectionStatus(csError)
'  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "GetStringResponse")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Unload frmMenu
'   End Select
'
'End Function
'
''---------------------------------------------------------------------
'Private Function GetByteResponse(ByVal vAddress As String, _
'                                 ByVal vTimeout As Integer) As Variant
''---------------------------------------------------------------------
''Called by:-    DownLoadMessages
''Using Microsofts Internet Control using HTTP this sub reads BINARY data from a
''TrialOffice server using GET and GetChunk
''---------------------------------------------------------------------
'
'Dim vtData As Variant ' Data variable.
'Dim vtResponseData As Variant
'Dim bDone As Boolean
'
'On Error GoTo ErrHandler
'
'    bDone = False
'
'    Inet1.RequestTimeout = vTimeout
'
'    '   ATN 14/2/2000
'    '   Need to set transfer properties each time the control is executed
'    SetTransferProperties vAddress
'    Inet1.Execute , "GET"
'    DoEvents
'    Inet1.Cancel
'    DoEvents
'    '   ATN 14/2/2000
'    SetTransferProperties vAddress
'    Inet1.Execute , "GET"
'    DoEvents
'
'    Do While Inet1.StillExecuting
'        DoEvents
'    Loop
'
'    vtData = Inet1.GetChunk(4096, icByteArray)
'
'    DoEvents
'    Do While Not bDone
'
'        vtResponseData = vtResponseData & vtData
'        ' Get next chunk.
'        vtData = Inet1.GetChunk(4096, icByteArray)
'        DoEvents
'        If Len(vtData) = 0 Then
'            bDone = True
'        End If
'    Loop
'
'    GetByteResponse = vtResponseData
'
'Exit Function
'ErrHandler:
'  Call SetConnectionStatus(csError)
'  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "GetByteResponse")
'        Case OnErrorAction.Ignore
'            Resume Next
'        Case OnErrorAction.Retry
'            Resume
'        Case OnErrorAction.QuitMACRO
'            Unload frmMenu
'   End Select
'
'End Function

'---------------------------------------------------------------------
Private Function DisplayTimeoutMessage()
'---------------------------------------------------------------------
' REVISIONS
' DPH 01/05/2002 - Changed display message
'---------------------------------------------------------------------

On Error GoTo ErrHandler
    
    If Me.Visible = True Then
    
        ' Error status
'        Call SetTextBoxDisplay("Couldn't connect to " & Me.MACROServerDesc & ".", trTransferFinished, False)
        Call SetTextBoxDisplay("The transfer connection to " & Me.MACROServerDesc & " has encountered problems." _
            & vbCrLf & vbTab & "Please check your connection to the MACRO server.", trTransferError)
        Call SetConnectionStatus(csError)
                
        cmdCancel.Visible = True
        cmdCancel.Caption = "&Close"
        
        Me.WindowState = vbNormal
        
        mbConnectionOK = False
        mbTransferOK = False
    End If
    
Exit Function
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "DisplayTimeoutMessage")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
   
End Function

'---------------------------------------------------------------------
Private Sub InitialiseHTTPControl()
'---------------------------------------------------------------------
' Set up the HTTP control for use within MACRO
'---------------------------------------------------------------------
' REVISIONS
'---------------------------------------------------------------------
On Error GoTo ErrHandler

    ' Use version 1.0
    Http1.Version = "HTTP/1.0"
    
    ' Don't Cache - probably get incorrect results
    Http1.Cache = False

    ' Set Timeout to 300 Seconds (will wait for page (in milliseconds))
    Http1.Timeout = 300000
        
    ' Proxy Settings
    If frmMenu.gTrialOffice.ProxyServer > "" Then
        Http1.Proxy = frmMenu.gTrialOffice.ProxyServer
        ' Http1.ProxyUsername
        ' Http1.ProxyPassword
    End If
    
    ' Set username / password in Method calls
    
    ' Security
    Http1.Security = httpAllowRedirectToHTTP + httpAllowRedirectToHTTPS
    
Exit Sub
ErrHandler:
  Call SetConnectionStatus(csError)
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "InitialiseHTTPControl")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
End Sub

'---------------------------------------------------------------------
Private Sub SetTransferPropertiesHTTP(ByVal sAddress As String, Optional nPortNumber As Integer = -1)
'---------------------------------------------------------------------
' Set HTTP URL with Port included
'---------------------------------------------------------------------

    If nPortNumber = -1 Then
        nPortNumber = frmMenu.gTrialOffice.PortNumber
    End If
    
    ' Port inclusion
    Http1.URL = AddPortToHTTPString(sAddress, nPortNumber)

End Sub

'---------------------------------------------------------------------
Private Sub SetTimeoutHTTP(ByVal nTimeout As Integer)
'---------------------------------------------------------------------
' Set Timeout on HTTP control
'---------------------------------------------------------------------

    ' Timeout in seconds (needs to be milliseconds)
    Http1.Timeout = CLng(nTimeout) * 1000
    
End Sub

'---------------------------------------------------------------------
Private Function GetStringResponseHTTP(ByVal vAddress As String, _
                                   ByVal vTimeout As Integer, _
                                   Optional sURLData As String = "") As String
'---------------------------------------------------------------------
'Called by:-    DownLoadMessages
'               DisplayGetMessages
'Using Dart Web Control using HTTP this sub reads from a
'TrialOffice server using POST
'---------------------------------------------------------------------
Dim sData As String
Dim sURLParamData As String

On Error GoTo ErrHandler
    
    sData = ""
    sURLParamData = ""
    If sURLData <> "" Then
        sURLParamData = sURLData
    End If
    
    ' Timeout
    Call SetTimeoutHTTP(vTimeout)
                    
    ' Set URL (with port)
    SetTransferPropertiesHTTP vAddress
    
    ' Post HTTP Request (will block wait with timeout)
    ' http1.post data(to be sent to server), request headers, response data,
    '               received headers, username, password
    Http1.Post sURLParamData, , sData, , msUser, msPassword
    
    ' Return collected string
    GetStringResponseHTTP = sData
    
Exit Function
ErrHandler:
  Call SetConnectionStatus(csError)
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "GetStringResponseHTTP")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
   
End Function

'---------------------------------------------------------------------
Private Function GetByteResponseHTTP(ByVal vAddress As String, _
                                 ByVal vTimeout As Integer) As Variant
'---------------------------------------------------------------------
'Called by:-    DownLoadMessages, DisplaySendMessages
'Using Darts Web Control using HTTP this sub reads BINARY data from a
'TrialOffice server using GET calls
'---------------------------------------------------------------------
Dim vData() As Byte

On Error GoTo ErrHandler
        
    ' Timeout in seconds (needs to be milliseconds)
    Call SetTimeoutHTTP(vTimeout)
                    
    ' Set URL (with port)
    SetTransferPropertiesHTTP vAddress
    
    ' Get HTTP Request (will block wait with timeout)
    ' http1.get data(to be sent to server), request headers, response data,
    '               received headers, username, password
    Http1.Get vData, , msUser, msPassword
    
    ' Return as variant
    GetByteResponseHTTP = vData
    
Exit Function
ErrHandler:
  Call SetConnectionStatus(csError)
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "GetByteResponseHTTP")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Unload frmMenu
   End Select
   
End Function

'---------------------------------------------------------------------
Private Function TestConnection(ByRef sErrorMsg As String) As Boolean
'---------------------------------------------------------------------
'REM 03/04/03
'Tests the connection to the server: Returns True if connection to server database successful,
' else returns False and appropriate error message
'---------------------------------------------------------------------
Dim sTestMessage As String
Dim vTestMsg As Variant
Dim sMessage As String
Dim bError As Boolean

    On Error GoTo ErrConnect
    
    bError = False
    
    'connect to test asp page
    sTestMessage = PostDataHTTP(Me.HTTPAddress & msTEST_CONNECTION_URL, mnDataTransferTimeout)
    
    vTestMsg = Split(sTestMessage, gsMSGSEPARATOR)
    
    If vTestMsg(0) = "Success" Then
        'connection the the server succeeded
        sMessage = "Connection to MACRO server successful." & vbCrLf
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
            sErrorMsg = sMessage
            TestConnection = False
        Else
            sErrorMsg = ""
            TestConnection = True
        End If
        
    Else
        sErrorMsg = "Connection to the MACRO server failed.  Please contact your system administrator."
        TestConnection = False
    End If
    
Exit Function
ErrConnect:
    sErrorMsg = "Connection to MACRO server failed. Please contact your system administrator."
    TestConnection = False
End Function


'---------------------------------------------------------------------
Private Function PostDataHTTP(ByVal vAddress As String, _
                                   ByVal vTimeout As Integer, _
                                   Optional sURLData As String = "", Optional nPortNumber As Integer = -1) As String
'---------------------------------------------------------------------
' Post Data using Dart Control catch error but raise
'---------------------------------------------------------------------
Dim sData As String
Dim sURLParamData As String

On Error GoTo ErrHandler
    
    sData = ""
    sURLParamData = ""
    If sURLData <> "" Then
        sURLParamData = sURLData
    End If
    
    ' Timeout
    Call SetTimeoutHTTP(vTimeout)
    
    ' Set URL (with port)
    SetTransferPropertiesHTTP vAddress, nPortNumber
    
    ' Post HTTP Request (will block wait with timeout)
    ' http1.post data(to be sent to server), request headers, response data,
    '               received headers, username, password
    Http1.Post sURLParamData, , sData, , msUser, msPassword
    
    ' Return collected string
    PostDataHTTP = sData
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|PostDataHTTP"
End Function

'---------------------------------------------------------------------
Private Function GetDataHTTP(ByVal vAddress As String, _
                                 ByVal vTimeout As Integer) As Variant
'---------------------------------------------------------------------
'Called by:-    DownLoadMessages
'Using Darts Web Control using HTTP this sub reads BINARY data from a
'TrialOffice server using GET calls
'---------------------------------------------------------------------
' REVISIONS
' DPH 02/05/2002 - Function as GetByteResponseHTTP except raises error
'                   rather than throws MACRO error
'---------------------------------------------------------------------
Dim vData() As Byte

On Error GoTo ErrHandler
        
    ' Timeout in seconds (needs to be milliseconds)
    Call SetTimeoutHTTP(vTimeout)
                    
    ' Set URL (with port)
    SetTransferPropertiesHTTP vAddress
    
    ' Get HTTP Request (will block wait with timeout)
    ' http1.get data(to be sent to server), request headers, response data,
    '               received headers, username, password
    Http1.Get vData, , msUser, msPassword
    
    ' Return as variant
    GetDataHTTP = vData
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|GetDataHTTP"
End Function

''---------------------------------------------------------------------
'Public Sub Register(ByVal vRegisterParameters As String, _
'                    ByVal vDisplay As Boolean, _
'                    ByRef rRegResult1 As String, _
'                    ByRef rRegResult2 As String)
''---------------------------------------------------------------------
'
'Dim msResult As String
'
'    If vDisplay Then
'        cmdCancel.Visible = False
'        txtMessage.Visible = False
'
'        lblStatus.Caption = "Connecting to " & Me.MACROServerDesc & "..."
'        Me.Caption = "Patient registration"
'        Me.Show
'    End If
'
'    On Error GoTo Timeout
'
'    msResult = GetStringResponse(Me.HTTPAddress & "register.asp?" & vRegisterParameters, mnDataTransferTimeout)
'
'    rRegResult1 = ExtractFirstItemFromList(msResult, ",")
'    rRegResult2 = msResult
'
'    If vDisplay Then
'        cmdCancel.Caption = "Close"
'        cmdCancel.Visible = True
''        cmdCancel.SetFocus
'
'        Select Case rRegResult1
'        Case "ERROR"
'            lblStatus.Caption = "Error code: " & rRegResult2
'            Me.Caption = "Patient could not be registered at " & Me.MACROServerDesc
'        Case "WARNING"
'            lblStatus.Caption = rRegResult2
'            Me.Caption = "Patient could not be registered at " & Me.MACROServerDesc
'        Case Else
'            lblStatus.Caption = "Subject Identification Code: " & rRegResult1
'            Me.Caption = "Patient registered at " & Me.MACROServerDesc
'        End Select
'
'        Screen.MousePointer = vbDefault
'        Me.WindowState = vbNormal
'    End If
'
'    'Me.Hide
'
'Exit Sub
'
'Timeout:
''   ATN 23/3/99
''   Set result to error if timeout happens
'    rRegResult1 = "ERROR"
'
'    If vDisplay Then
'        DisplayTimeoutMessage
'    Else
'        End
'    End If
'
'End Sub


''---------------------------------------------------------------------
'Public Sub Randomise(ByVal vRandomisationParameters As String, _
'                         ByVal vDisplay As Boolean, _
'                         ByRef rRandResult1 As String, _
'                         ByRef rRandResult2 As String)
''---------------------------------------------------------------------
'
'Dim msResult As String
'
'    If vDisplay Then
'        cmdCancel.Visible = False
'
'        txtMessage.Visible = False
'
'        lblStatus.Caption = "Connecting to " & Me.MACROServerDesc & "..."
'        Me.Caption = "Patient randomisation"
'        Me.Show
'    End If
'
'    On Error GoTo Timeout
'
'    msResult = GetStringResponse(Me.HTTPAddress & "random.asp?" & vRandomisationParameters, mnDataTransferTimeout)
'
'    rRandResult1 = ExtractFirstItemFromList(msResult, ",")
'    rRandResult2 = msResult
'
'    If vDisplay Then
'        cmdCancel.Caption = "Close"
'        cmdCancel.Visible = True
''        cmdCancel.SetFocus
'
'        Select Case rRandResult1
'        Case "ERROR"
'            lblStatus.Caption = "Error code: " & rRandResult2
'            Me.Caption = "Patient could not be randomised at " & Me.MACROServerDesc
'        Case "WARNING"
'            lblStatus.Caption = rRandResult2
'            Me.Caption = "Patient could not be randomised at " & Me.MACROServerDesc
'        Case Else
'            lblStatus.Caption = "Treatment: " & rRandResult1
'            Me.Caption = "Patient randomised at " & Me.MACROServerDesc
'        End Select
'
'        Screen.MousePointer = vbDefault
'        Me.WindowState = vbNormal
'    End If
'    'Me.Hide
'Exit Sub
'
'Timeout:
''   ATN 23/3/99
''   Set result to error if timeout happens
'    rRandResult1 = "ERROR"
'
'    If vDisplay Then
'        DisplayTimeoutMessage
'    Else
'        End
'    End If
'
'End Sub

'---------------------------------------------------------------------
Public Function TrialOfficeRegistration(ByVal sTrialName As String, _
                                    ByVal sSite As String, _
                                    ByVal lPersonId As Long, _
                                    ByVal sPrefix As String, _
                                    ByVal sSuffix As String, _
                                    ByVal nUsePrefix As Integer, _
                                    ByVal nUseSuffix As Integer, _
                                    ByVal lStartNumber As Long, _
                                    ByVal nNumberWidth As Integer, _
                                    ByVal sUCheckString As String, _
                                    ByRef sSubjectIdentifier As String) As Integer
'---------------------------------------------------------------------
' Mo Morris (& NCJ) 24 Nov 2000
' Make a registration request to Trial Office
' DPH 18/04/2002 - Upgraded to use new control
'---------------------------------------------------------------------
Dim sResultOfRegistrationCall As String
Dim sHTTPAddress As String
Dim sPrompt As String

    On Error GoTo Timeout
        
    ' Get the Trial Office URL
    sHTTPAddress = frmMenu.gTrialOffice.HTTPAddress
    'changed Mo Morris 26/2/01, URLCharToHexEncoding now called due to the possibiliy of special
    'characters in sPrefix, sSuffix or UCheckString
    If sHTTPAddress > "" Then
        sResultOfRegistrationCall = PostDataHTTP(sHTTPAddress & msRegistrationURL & "?TrialName=" & sTrialName _
                                                    & "&Site=" & sSite _
                                                    & "&PersonId=" & lPersonId _
                                                    & "&Prefix=" & URLCharToHexEncoding(sPrefix) _
                                                    & "&Suffix=" & URLCharToHexEncoding(sSuffix) _
                                                    & "&UsePrefix=" & nUsePrefix _
                                                    & "&UseSuffix=" & nUseSuffix _
                                                    & "&StartNumber=" & lStartNumber _
                                                    & "&NumberWidth=" & nNumberWidth _
                                                    & "&UCheckString=" & URLCharToHexEncoding(sUCheckString), mnDataTransferTimeout)
    Else
        ' No HTTP address
        sPrompt = "There are no Trial Office Server connection details set up"
        Call DialogError(sPrompt)
        sSubjectIdentifier = ""
        TrialOfficeRegistration = eRegResult.RegError
        Exit Function
    End If
    
    On Error GoTo ErrHandler
    
    'The Registration server returns a string in the format SUBJECTIDENTIFIER<br>RESULTCODE
    'The SubjectIdentifier is returned to the calling procedure via the ByRef variable sSubjectIdentifier
    'The result code is a general status code and is return as the result of this function
    If Left$(sResultOfRegistrationCall, 6) = "<html>" Then
        sPrompt = "The Trial Office Server is not responding correctly. " & vbNewLine _
                & "It returned the following error:-" _
                & vbNewLine & sResultOfRegistrationCall
        Call DialogError(sPrompt)
        sSubjectIdentifier = ""
        TrialOfficeRegistration = eRegResult.RegError
    Else
        sSubjectIdentifier = ExtractFirstItemFromList(sResultOfRegistrationCall, "<br>")
        ' What's left should be the numeric result code
        If IsNumeric(sResultOfRegistrationCall) Then
            TrialOfficeRegistration = CInt(sResultOfRegistrationCall)
        Else
            TrialOfficeRegistration = eRegResult.RegError
        End If
    End If

Exit Function

Timeout:
    ' Connection timed out before response received
        sPrompt = "Could not connect to MACRO Server"
        Call DialogError(sPrompt)
        sSubjectIdentifier = ""
        TrialOfficeRegistration = eRegResult.RegError
Exit Function

ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "TrialOfficeRegistration")
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
Public Sub CollectInfoAboutUpload()
'---------------------------------------------------------------------
' Collect information about the amount of data to be transferred
'---------------------------------------------------------------------
Dim rsDownloadData As ADODB.Recordset
Dim sSQL As String
Dim nLogDetails As Integer
Dim nLoginLog As Integer

On Error GoTo ErrHandler
    
    ' Data to send from site to server
    ' count the System Messages
    mlNoSysMessagesUpload = 0
    sSQL = "SELECT COUNT(*) FROM Message" _
        & " WHERE TrialSite = '" & Me.Site & "'" _
        & " AND MessageReceived = " & MessageReceived.NotYetReceived _
        & " AND MessageDirection = " & MessageDirection.MessageIn
    Set rsDownloadData = New ADODB.Recordset
    rsDownloadData.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not rsDownloadData.EOF Then
        If Not IsNull(rsDownloadData(0).Value) Then
            mlNoSysMessagesUpload = rsDownloadData(0).Value
        End If
    End If
    rsDownloadData.Close
    
    'count the logdetails messages that will be added to the message table during datatransfer
    sSQL = "SELECT COUNT(*) FROM LogDetails" _
        & " WHERE Status = 0"
    Set rsDownloadData = New ADODB.Recordset
    rsDownloadData.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not rsDownloadData.EOF Then
        If Not IsNull(rsDownloadData(0).Value) Then
            'calc the number of LogDetail messages that will be created, always need to round up, so use negative integer of negative
            nLogDetails = -Int(-(rsDownloadData(0).Value / 20))
            
            mlNoSysMessagesUpload = mlNoSysMessagesUpload + nLogDetails
        End If
    End If
    rsDownloadData.Close
    
    'count the userlog and messages that will be added to the message table during datatransfer
    sSQL = "SELECT COUNT(*) FROM LoginLog" _
        & " WHERE Status = 0"
    Set rsDownloadData = New ADODB.Recordset
    rsDownloadData.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not rsDownloadData.EOF Then
        If Not IsNull(rsDownloadData(0).Value) Then
            'calc the number of LogDetail messages that will be created, always need to round up, so use negative integer of negative
            nLoginLog = -Int(-(rsDownloadData(0).Value / 20))
            
            mlNoSysMessagesUpload = mlNoSysMessagesUpload + nLoginLog
        End If
    End If
    rsDownloadData.Close
    
    ' count subjects to send
    mlNoTransSubjects = 0
    sSQL = "SELECT count(*) FROM TrialSubject WHERE Changed = " & Changed.Changed
    Set rsDownloadData = New ADODB.Recordset
    rsDownloadData.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not rsDownloadData.EOF Then
        If Not IsNull(rsDownloadData(0).Value) Then
            mlNoTransSubjects = rsDownloadData(0).Value
        End If
    End If
    rsDownloadData.Close
    
    ' Count MIMessages to Send
    mlNoTransMIMessageUp = 0
    sSQL = "SELECT count(*) FROM MIMessage " _
    & "WHERE MIMessageSource = " & TypeOfInstallation.RemoteSite _
    & " AND MIMessageSent = 0"
    rsDownloadData.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not rsDownloadData.EOF Then
        If Not IsNull(rsDownloadData(0).Value) Then
            mlNoTransMIMessageUp = rsDownloadData(0).Value
        End If
    End If
    rsDownloadData.Close
    
    ' count laboratory to send
    mlNoTransLaboratoryUp = 0
    sSQL = "SELECT count(*) FROM Laboratory WHERE Changed = " & Changed.Changed
    rsDownloadData.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not rsDownloadData.EOF Then
        If Not IsNull(rsDownloadData(0).Value) Then
            mlNoTransLaboratoryUp = rsDownloadData(0).Value
        End If
    End If
    rsDownloadData.Close
    
    Set rsDownloadData = Nothing
Exit Sub

ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CollectInfoAboutDownload", "frmDataTransfer")
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
Public Function CollectInfoAboutDownload() As Boolean
'---------------------------------------------------------------------
' Collect information about the amount of data to be transferred
' from the strings passed from the server
' bMessage TRUE - Message table FALSE - MIMessage Table
'REM 12/12/02 - added mlNoSystemMessages to calc the number of system messages
'---------------------------------------------------------------------
Dim sServerReply As String
Dim sHTTPAddress As String
Dim sMessages() As String
Dim i As Integer
Dim nPos As Integer
Dim bOK As Boolean

On Error GoTo ErrHandler

    ' Connect to server & get data from exchange_download_info.asp
    sHTTPAddress = Me.HTTPAddress & msDownloadInfoURL & "?site=" & Me.Site

    On Error GoTo ErrWebConn
    sServerReply = PostDataHTTP(sHTTPAddress, mnDataTransferTimeout)
    
    On Error GoTo ErrHandler

    ' Initialise
    mlNoTransLockFreeze = 0
    mlNoTransStudyUpdate = 0
    mlNoTransStudyStatus = 0
    mlNoTransLaboratoryDown = 0
    mlNoTransMIMessageDown = 0
    mlNoSysMessagesDownload = 0
    mlNoTransReportFiles = 1
    bOK = True
    
    ' if contains <html> then error
    If sServerReply = "" Or InStr(1, LCase(sServerReply), "<html") > 0 _
        Or InStr(1, sServerReply, "<font") > 0 Or Left(sServerReply, 5) = "ERROR" Then
        bOK = False
    End If
    
    If bOK Then
        sMessages() = Split(sServerReply, vbCrLf)
        For i = 0 To UBound(sMessages()) - 1
            Select Case Left(sMessages(i), 6)
                Case "No stu"
                    ' No study messages
                Case "A new "
                    ' Study definition
                    mlNoTransStudyUpdate = mlNoTransStudyUpdate + 1
                Case "Lockin"
                    ' Locking
                    mlNoTransLockFreeze = mlNoTransLockFreeze + 1
                Case "Freezi"
                    ' Freezing
                    mlNoTransLockFreeze = mlNoTransLockFreeze + 1
                Case "Study "
                    ' study status
                    mlNoTransStudyStatus = mlNoTransStudyStatus + 1
                Case "Labora"
                    ' Laboratory
                    mlNoTransLaboratoryDown = mlNoTransLaboratoryDown + 1
                Case "MIMess"
                   ' MIMessages rolled into one
                   mlNoTransMIMessageDown = CInt(Right(sMessages(i), Len(sMessages(i)) - 11))
                Case Else ' this will take care of all system messages
                   mlNoSysMessagesDownload = mlNoSysMessagesDownload + 1
            End Select
        Next
    End If
    
    CollectInfoAboutDownload = bOK
Exit Function

ErrWebConn:
    CollectInfoAboutDownload = False
    Exit Function
    
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CollectInfoAboutDownload", "frmDataTransfer")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
   End Select
End Function

' DPH 16/05/2002 - Added command line data transfer with conditional compilation arguements
#If BackgroundXfer Then

Public Sub BackGroundDisplay()
    ' Display form modeless
    Me.Show vbModeless
        
    ' Hide buttons
    Me.cmdCancel.Visible = False
    Me.cmdStart.Visible = False
    
    ' Display Command line data transfer message
    txtMessage.Text = "Command line data transfer started." & vbCrLf
    
    ' Run Transfer
    mbTransferOK = True

    ' Get messages / data from server
    If DisplayGetMessages Then
        ' Send data from the site to the server
        DisplaySendMessages
    End If

    'REM 04/12/02 - added system message transfer
    'send system messages from site to server
    If DisplaySendSystemMessages Then
        'get system messages from server
        DisplayGetSystemMessages
    End If

End Sub


#End If


'---------------------------------------------------------------------
Private Function DownLoadLFMessages() As Boolean
'---------------------------------------------------------------------
' NCJ 19 Dec 02 - Download Lock/Freeze messages
' NCJ 20 Jan 03 - Make sure we lock the subject when executing an LF message
'---------------------------------------------------------------------
Dim sNextMessage As String
Dim sSQL As String
Dim rsTemp As ADODB.Recordset
' Used to control percentage increment
Dim nPercentInc As Integer

Dim lErrNo As Long
Dim sErrDesc As String

' The Lock Freeze Business objects
Dim oLFMsg As LFMessage
Dim oExistingLFMsg As LFMessage
Dim oFlocker As LockFreeze

Dim oTimezone As TimeZone
Dim lLockedStudyId As Long
Dim lLockedSubjId As Long
Dim sLockToken As String

    On Error GoTo ErrHandler
    
    Set oTimezone = New TimeZone
    Set oExchange = New clsExchange
    
    ' These values keep track of which subject we've locked
    lLockedStudyId = 0
    lLockedSubjId = 0
    sLockToken = ""     ' The lock token for when we lock subjects
    
    Me.cmdCancel.Visible = True
    Call SetTextBoxDisplay("Downloading Lock/Freeze messages from " & Me.MACROServerDesc, trTransferData)
    nPercentInc = 1
    
    On Error GoTo Timeout
    
    sNextMessage = PostDataHTTP(Me.HTTPAddress & msGetNextLFMessageURL & "?site=" & Me.Site, mnDataTransferTimeout)
    
    On Error GoTo ErrHandler

    ' DPH 01/05/2002 - Pick up if an error has occurred in the ASP page
    If Left(sNextMessage, 5) = "ERROR" Or InStr(1, LCase(sNextMessage), "<html") > 0 _
        Or (Len(sNextMessage) > 1 And InStr(1, sNextMessage, "<br>") = 0) Then
        Call SetTextBoxDisplay("Lock/Freeze messages could not be downloaded", trStepCompleted, False)
        ' log error
        gLog gsDOWNLOAD_LFMESG, "Lock/Freeze messages could not be downloaded."
        mbTransferOK = False
        Exit Function
    End If
    
    Set oFlocker = New LockFreeze
    
    If sNextMessage <> "." Then
        ' start timing
        Call mDataTransferTime.StartTiming(LFMessagesDownSection)
    End If
    
    ' Unwrap into an LFMessage object
    Do While sNextMessage <> "."
        ' NCJ 10 June 04 - Bug 2296 - Must reset oLFMsg each time round the loop
        ' to avoid the TrialId being incorrectly reused!
        Set oLFMsg = Nothing
        Set oLFMsg = New LFMessage
        With oLFMsg
            'extract the individual fields from the user message
            .MessageId = CLng(ExtractFirstItemFromList(sNextMessage, "<br>"))
            .StudyName = ExtractFirstItemFromList(sNextMessage, "<br>")
            .Site = ExtractFirstItemFromList(sNextMessage, "<br>")
            .SubjectId = CLng(ExtractFirstItemFromList(sNextMessage, "<br>"))
            .VisitId = CLng(ExtractFirstItemFromList(sNextMessage, "<br>"))
            .VisitCycle = CInt(ExtractFirstItemFromList(sNextMessage, "<br>"))
            .EFormId = CLng(ExtractFirstItemFromList(sNextMessage, "<br>"))
            .EFormCycle = CInt(ExtractFirstItemFromList(sNextMessage, "<br>"))
            .ResponseId = CLng(ExtractFirstItemFromList(sNextMessage, "<br>"))
            .ResponseCycle = CInt(ExtractFirstItemFromList(sNextMessage, "<br>"))
            .QuestionId = CLng(ExtractFirstItemFromList(sNextMessage, "<br>"))
            .MsgSource = CInt(ExtractFirstItemFromList(sNextMessage, "<br>"))
            .ActionType = CInt(ExtractFirstItemFromList(sNextMessage, "<br>"))
            .Scope = CInt(ExtractFirstItemFromList(sNextMessage, "<br>"))
            .UserName = ExtractFirstItemFromList(sNextMessage, "<br>")
            .UserNameFull = ExtractFirstItemFromList(sNextMessage, "<br>")
            .RollbackSource = CInt(ExtractFirstItemFromList(sNextMessage, "<br>"))
            .RollbackMsgId = CLng(ExtractFirstItemFromList(sNextMessage, "<br>"))
            .MsgCreatedTimestamp = CDbl(StandardNumToLocal(ExtractFirstItemFromList(sNextMessage, "<br>")))
            .MsgCreatedTimestamp_TZ = CInt(ExtractFirstItemFromList(sNextMessage, "<br>"))
            .SentTimestamp = CDbl(StandardNumToLocal(ExtractFirstItemFromList(sNextMessage, "<br>")))
            .SentTimestamp_TZ = CInt(ExtractFirstItemFromList(sNextMessage, "<br>"))
            ' Store date of receipt
            .ReceivedTimestamp = IMedNow
            .ReceivedTimestamp_TZ = oTimezone.TimezoneOffset
            ' Not processed yet
            .ProcessedStatus = LFProcessStatus.lfpUnProcessed
            .ProcessedTimestamp = 0
            .ProcessedTimestamp_TZ = 0
            
            ' Check we haven't had this one already
            Set oExistingLFMsg = New LFMessage
            If oExistingLFMsg.Load(MacroADODBConnection, .Site, .MsgSource, .MessageId) Then
                ' We've had this one before so ignore it
            Else
                Call .Save(MacroADODBConnection)
                ' NB Save will have filled in the correct StudyId
                ' Now do it if possible, or create a rollback message if not
                If .CanExecute(MacroADODBConnection) Then
                    ' Theoretically we can do this message
                    ' If a new subject, need to unlock previous one
                    If (.SubjectId <> lLockedSubjId) Or (.StudyId <> lLockedStudyId) Then
                        ' New subject - need to lock it
                        If sLockToken > "" Then
                            ' Unlock previous subject
                            Call oExchange.RemoveSubjectLock(lLockedStudyId, Me.Site, lLockedSubjId, sLockToken)
                        End If
                        sLockToken = oExchange.GetSubjectLock(goUser.UserName, .StudyId, .Site, .SubjectId)
                        lLockedSubjId = .SubjectId
                        lLockedStudyId = .StudyId
                    End If
                    If sLockToken > "" Then
                        ' We got the lock on the subject
                        Call .DoAction(MacroADODBConnection, True)
                    Else
                        ' Someone's meddling with this subject - refuse the message
                        Call oFlocker.RefuseMessage(MacroADODBConnection, oLFMsg)
                    End If
                Else
                    ' Can't do the message (e.g. because of changed data)
                    Call oFlocker.RefuseMessage(MacroADODBConnection, oLFMsg)
                End If
            End If
            
            On Error GoTo Timeout
                    
            'Note that we pass the LFMessageID that has just been
            'successfully processed, for the purpose of setting it to sent (together with its sent time)
            sNextMessage = PostDataHTTP(Me.HTTPAddress & msGetNextLFMessageURL _
                            & "?site=" & Me.Site _
                            & "&PreviousMessageID=" & .MessageId _
                            & "&PreviousMessageSent=" & ConvertLocalNumToStandard(CStr(.SentTimestamp)) _
                            & "&PreviousMessageSentTZ=" & .SentTimestamp_TZ, mnDataTransferTimeout)
            
            On Error GoTo ErrHandler
            
            ' DPH 01/05/2002 - Improve error handling
            If Left(sNextMessage, 5) = "ERROR" Or InStr(1, LCase(sNextMessage), "<html") > 0 _
            Or (Len(sNextMessage) > 1 And InStr(1, sNextMessage, "<br>") = 0) Then
                Call SetTextBoxDisplay("Some Lock/Freeze messages could not be downloaded", trStepCompleted, False)
                ' log error
                gLog gsDOWNLOAD_LFMESG, "Some Lock/Freeze messages could not be downloaded."
                mbTransferOK = False
                If sLockToken > "" Then
                    ' Unlock last locked subject
                    Call oExchange.RemoveSubjectLock(lLockedStudyId, Me.Site, lLockedSubjId, sLockToken)
                End If
                Exit Function
            End If
    
            ' Increment progress bar / mimessage timer
            Call mDataTransferTime.IncrementLFMessagesDown
            If (prgProgress.Value + nPercentInc) < 20 Then
                Call SetProgressBar(prgProgress.Value + nPercentInc, "")
            End If
        End With
    Loop
        
    ' Unlock last locked subject
    If sLockToken > "" Then
        Call oExchange.RemoveSubjectLock(lLockedStudyId, Me.Site, lLockedSubjId, sLockToken)
    End If
    
    If Me.Visible = True Then
        Me.cmdCancel.Caption = "&Cancel"
        Me.cmdCancel.Visible = True
        Call SetTextBoxDisplay("All Lock/Freeze messages from " & Me.MACROServerDesc & " downloaded", trStepCompleted)
    End If
        
    Set oTimezone = Nothing
    Set oFlocker = Nothing
    
    Me.MousePointer = vbDefault
        
    DownLoadLFMessages = True
    
Exit Function

Timeout:
    'DisplayTimeoutMessage
    DownLoadLFMessages = False
    If sLockToken > "" Then
        ' Unlock last locked subject
        Call oExchange.RemoveSubjectLock(lLockedStudyId, Me.Site, lLockedSubjId, sLockToken)
    End If
Exit Function

ErrHandler:
    mbTransferOK = False
    Call SetTextBoxDisplay("An error was encountered while dowloading Lock/Freeze messages" & vbCrLf & "Error Description: " & Err.Description & "Error Number: " & Err.Number, trStepCompleted, mbTransferOK)
    If sLockToken > "" Then
        ' Unlock last locked subject
        Call oExchange.RemoveSubjectLock(lLockedStudyId, Me.Site, lLockedSubjId, sLockToken)
    End If
    DownLoadLFMessages = True 'set to true so Data Transfer can continue
End Function

'---------------------------------------------------------------------
Public Function SendLFMessages() As Boolean
'---------------------------------------------------------------------
' Send all the Site's unsent LFMessages to the Server
'REM 08/04/03 - changed to a function
'---------------------------------------------------------------------
Dim sHTTPAddress As String
Dim sHTTPData As String
Dim oTimezone As TimeZone
Dim colUnsentMsgs As Collection
Dim oFlocker As LockFreeze
Dim oLFMsg As LFMessage
Dim dblSentTime As Double
Dim sMessageInfo As String
Dim sResultOfPosting As String
Dim i As Long
Dim nPercentInc As Integer

    On Error GoTo ErrHandler
    
    Me.cmdCancel.Visible = True

    Set oTimezone = New TimeZone
    Set oFlocker = New LockFreeze
    
    Set colUnsentMsgs = oFlocker.GetMessagesToTransfer(MacroADODBConnection, TypeOfInstallation.RemoteSite, Me.Site)
    
    If colUnsentMsgs.Count = 0 Then
        Call SetTextBoxDisplay("No Lock/Freeze messages to be sent to " & Me.MACROServerDesc, trTransferData)
    Else
        Call SetTextBoxDisplay("Sending Lock/Freeze messages to " & Me.MACROServerDesc, trTransferData)
        Call SetTextBoxDisplay(colUnsentMsgs.Count & " Lock/Freeze messages", trTransferData)
        ' start timing
        Call mDataTransferTime.StartTiming(LFMessagesUpSection)
    
        For i = 1 To colUnsentMsgs.Count
            ' calculate percentage to increase each LFmessage
            ' sending LFmessage 15%
            nPercentInc = CInt(15 \ colUnsentMsgs.Count)
            If nPercentInc = 0 Then
                nPercentInc = 1
            End If
            
            On Error GoTo Timeout
            
            'store the time sent for the purpose of using the same time when updating it in the
            'sending Client database
            dblSentTime = IMedNow
            
            sHTTPAddress = Me.HTTPAddress & msReceiveLFMessageURL
            
            Set oLFMsg = colUnsentMsgs(i)
            With oLFMsg
                sHTTPData = "ID=" & .MessageId _
                    & "&TrialName=" & .StudyName & "&Site=" & .Site & "&PersonID=" & .SubjectId _
                    & "&VisitID=" & .VisitId & "&VisitCycle=" & .VisitCycle _
                    & "&CRFPageID=" & .EFormId & "&CRFPageCycle=" & .EFormCycle _
                    & "&ResponseTaskID=" & .ResponseId & "&ResponseCycle=" & .ResponseCycle _
                    & "&QuestionID=" & .QuestionId _
                    & "&Source=" & .MsgSource & "&ActionType=" & .ActionType & "&Scope=" & .Scope _
                    & "&UserName=" & URLCharToHexEncoding(.UserName) & "&UserNameFull=" & URLCharToHexEncoding(.UserNameFull) _
                    & "&RollbackSource=" & .RollbackSource & "&RollbackMessageID=" & .RollbackMsgId _
                    & "&Processed=" & .ProcessedStatus _
                    & "&MsgTimeStamp=" & LocalNumToStandard(CStr(.MsgCreatedTimestamp)) _
                    & "&MsgTimeStampTZ=" & .MsgCreatedTimestamp_TZ _
                    & "&ProcessTimeStamp=" & LocalNumToStandard(CStr(.MsgCreatedTimestamp)) _
                    & "&ProcessTimeStampTZ=" & .MsgCreatedTimestamp_TZ _
                    & "&SentTimeStamp=" & LocalNumToStandard(dblSentTime) _
                    & "&SentTimeStampTZ=" & oTimezone.TimezoneOffset
            End With
            
            sResultOfPosting = PostDataHTTP(sHTTPAddress, 90, sHTTPData)
            
            DoEvents
            
            On Error GoTo ErrHandler
            
            If sResultOfPosting = "SUCCESS" Then
                'set MessageSent = now for the message that has just been successfully sent to the server
                Call oLFMsg.SetAsSent(MacroADODBConnection)
                sMessageInfo = oLFMsg.ActionText & " " & oLFMsg.ScopeText
                sMessageInfo = sMessageInfo & " for Subject " & oLFMsg.Site & "/" & oLFMsg.SubjectId
                sMessageInfo = sMessageInfo & " sent to " & Me.MACROServerDesc & " successfully"
                Call SetTextBoxDisplay(sMessageInfo, trTransferData)
            Else
                Call SetTextBoxDisplay("Lock/Freeze messages could not be sent to " & Me.MACROServerDesc, trTransferData)
                mbTransferOK = False
                'the last message was not read successfully so exit the loop by moving past the last record
                i = colUnsentMsgs.Count + 1
            End If
            ' update timing class
            Call mDataTransferTime.IncrementLFMessagesUp
            ' Increment progress bar
            If (prgProgress.Value + nPercentInc) < 68 Then
                Call SetProgressBar(prgProgress.Value + nPercentInc, "")
            End If
            
        Next i
        
        ' NCJ 14 Jan 03 - We no longer do this here (it's done in AutoImport)
'        If mbTransferOK Then
'            ' NCJ 7 Jan 03 - Must get the Server to process the Rollbacks just sent (if any)
'            sHTTPAddress = Me.HTTPAddress & msProcessLFMessagesURL
'            sHTTPData = "Site=" & Me.Site
'            sResultOfPosting = PostDataHTTP(sHTTPAddress, 90, sHTTPData)
'            If sResultOfPosting = "SUCCESS" Then
'                Call SetTextBoxDisplay("Lock/Freeze messages have been processed on " & Me.MACROServerDesc, trStepCompleted)
'            Else
'                mbTransferOK = False
'                Call SetTextBoxDisplay("Lock/Freeze messages could not be processed on " & Me.MACROServerDesc, trTransferData)
'            End If
'        End If

    End If
    
    Set oTimezone = Nothing
    Set oFlocker = Nothing
    
    SendLFMessages = True
    
Exit Function
Timeout:
    DisplayTimeoutMessage
    SendLFMessages = False

Exit Function
ErrHandler:
  mbTransferOK = False
  Call SetTextBoxDisplay("An error was encountered while sending Lock/Freeze messages to the server" & vbCrLf & "Error Description: " & Err.Description & "Error Number: " & Err.Number, trStepCompleted, mbTransferOK)
  SendLFMessages = True 'set to true as can still continue with the rest of data transfer

End Function

'---------------------------------------------------------------------
Public Sub DisplayGetModifiedReports()
'---------------------------------------------------------------------
' Get the modified reports from the server
'---------------------------------------------------------------------
'REVISIONS:
'   REM 13/10/03 - Added GetLastTransferDate routine to retrieve last transfer date from the message table
'---------------------------------------------------------------------
Dim sReportsMessage As String
Dim oModReports As New SysDataXfer
Dim sHTTPAddress As String
Dim sHTTPData As String
Dim sSplitMsg() As String
Dim sErrMsg As String
Dim sZIPfile As String
Dim sZipFileAndPath As String
Dim vSaveZipErr As Variant
Dim mbReport() As Byte
Dim mnFreeFile As Integer
Dim bSucceeded As Boolean
Dim sConfirmMessage As String
Dim sSucceeded As String
Dim sLastTransDate As String
Dim bSuccess As Boolean

    On Error GoTo Errlabel

    cmdCancel.Visible = True
    
    'boolean to set if report transfer was success
    bSuccess = True
    
    'return the last time data transfer was carried out
    sLastTransDate = GetLastTransferDate(sErrMsg)

    sHTTPAddress = Me.HTTPAddress & msGetModifiedReportFiles
    sHTTPData = "username=" & goUser.UserName & "&site=" & Me.Site & "&lasttrans=" & sLastTransDate
    
    Call SetConnectionStatus(csDownloadingReportFiles)
    Call SetTextBoxDisplay("Checking for report files on server " & Me.MACROServerDesc & "...", trStepStarted)
    
    'check for err message from routine GetLastTransferDate
    If sErrMsg <> "" Then 'if err msg then will download reports anyway
        Call SetTextBoxDisplay("All report files will be downloaded because last transfer date could not be determined!", trTransferData)
        'Replace invalid chars before writing to the sys log
        sErrMsg = ReplaceInvalidCharsForDataXfer(sErrMsg)
        Call gLog(gsREPORT_XFER_ERR, sErrMsg)
        'set success to flase so display shows error icon
        bSuccess = False
    End If

    If mbCancelled Then
        Exit Sub
    End If
    
    Call mDataTransferTime.StartTiming(ReportFilesSection)
    
    On Error GoTo Timeout
    'get reports from the server
    sReportsMessage = PostDataHTTP(sHTTPAddress, mnDataTransferTimeout, sHTTPData)
    
    On Error GoTo Errlabel
    
    If sReportsMessage = "" Then
        GoTo Timeout
    ElseIf InStr(1, LCase(sReportsMessage), "<html") > 0 Then
        GoTo ASPError
    Else
        ' split received message and check for errors
        sSplitMsg = Split(sReportsMessage, "<br>")
        sZIPfile = sSplitMsg(0)
        sErrMsg = ""
        If UBound(sSplitMsg) >= 1 Then
            sErrMsg = sSplitMsg(1)
        End If
        'Error
        If sErrMsg <> "" Then GoTo Errlabel
        
        If sZIPfile <> "" Then
            ' get zip file from server
            Call SetTextBoxDisplay("Downloading report ZIP file from " & Me.MACROServerDesc & "...", trTransferData)
            sZipFileAndPath = gsIN_FOLDER_LOCATION & sZIPfile
            mbReport() = GetDataHTTP(Me.HTTPAddress & sZIPfile, 300)
            mnFreeFile = FreeFile
            Open sZipFileAndPath For Binary Access Write As #mnFreeFile
            Put #mnFreeFile, , mbReport()
            Close #mnFreeFile
            
            vSaveZipErr = ""
            ' save files to reports folder using Transfer DLL
            If oModReports.WriteReportFiles(goUser.DatabaseCode, sZipFileAndPath, vSaveZipErr, _
                                    goUser.UserName, Me.Site) Then
                ' Log download of reports as succeeded
                Call SetTextBoxDisplay("Downloaded reports for Site " & Me.Site & " successfully", trTransferData)
                gLog gsREPORT_XFER_SITE, "Downloaded reports for Site " & Me.Site & " successfully"
                
                ' confirm receipt of files
                sHTTPData = sHTTPData & "&confirm=yes"
                sConfirmMessage = PostDataHTTP(sHTTPAddress, mnDataTransferTimeout, sHTTPData)
            
                If sConfirmMessage = "" Then
                    GoTo Timeout
                ElseIf InStr(1, LCase(sConfirmMessage), "<html") > 0 Then
                    GoTo ASPError
                Else
                    ' split received message and check for errors
                    sSplitMsg = Split(sConfirmMessage, "<br>")
                    sSucceeded = sSplitMsg(0)
                    sErrMsg = ""
                    If UBound(sSplitMsg) >= 1 Then
                        sErrMsg = sSplitMsg(1)
                    End If
                    
                    If sSucceeded <> "SUCCESS" Then
                        sErrMsg = "Unknown Error occurred attempting to confirm receipt of reports files "
                    End If
                   
                    'Error
                    If sErrMsg <> "" Then GoTo Errlabel
                End If
            Else
                ' Failure in storing zip
                If CStr(vSaveZipErr) <> "" Then
                    sErrMsg = CStr(vSaveZipErr)
                Else
                    sErrMsg = ""
                End If
                ' Log download of reports as failed
                Call SetTextBoxDisplay("Download of reports for Site " & Me.Site & " failed", trTransferData)
                gLog gsREPORT_XFER_SITE, "Download of reports for Site " & Me.Site & " failed"
                'Error
                If sErrMsg <> "" Then GoTo Errlabel
            End If
        Else
            Call SetTextBoxDisplay("No modified reports to download for Site " & Me.Site & " ", trTransferData)
        End If
    End If
    
    Call mDataTransferTime.StopTiming(ReportFilesSection)
    
    ' Set Progress Bar to 72% (a guess)
    Call SetProgressBar(72, "")

    Call SetTextBoxDisplay("Completed reports download from " & Me.MACROServerDesc, trStepCompleted, bSuccess)

    If mbCancelled Then
        'Call SetTextBoxDisplay("Transfer cancelled", trTransferCancelled)
        Exit Sub
    End If
    
Exit Sub
Timeout:
    DisplayTimeoutMessage

Exit Sub
ASPError:
    Call SetConnectionStatus(csError)
    Call SetTextBoxDisplay("The MACRO server is not responding correctly. Please contact your system administrator.", trTransferError)
    'Call DialogError("The MACRO server is not responding correctly. Please contact your system administrator." _
        , "Data Transfer Call Error")
    mbConnectionOK = False
    cmdCancel.Caption = "&Close"
    
Exit Sub
Errlabel:
    If sErrMsg = "" Then
        sErrMsg = "Error Description: " & Err.Description & ", Error Number: " & Err.Number
    End If
    Call SetTextBoxDisplay("System report download error, " & sErrMsg, trStepCompleted, False)
    'REM 27/03/03 - Replace invalid chars before writing to the sys log
    sErrMsg = ReplaceInvalidCharsForDataXfer(sErrMsg)
    Call gLog(gsREPORT_XFER_ERR, sErrMsg)
    mbTransferOK = False
End Sub

'---------------------------------------------------------------------
Private Function GetLastTransferDate(ByRef sErrMsg As String) As String
'---------------------------------------------------------------------
'REM 13/10/03
'Returns the date the last time Data Transfer was performed
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsMessages As ADODB.Recordset
Dim sDate As String
Dim dDate As Double
    
    On Error GoTo Errlabel
    
    'get the timestamp from all System Messages (use these as every data transfer sends system messages)
    'order them by Descending date so we get most rescent first
    sSQL = "SELECT MessageTimeStamp, MessageDirection "
    sSQL = sSQL & "MessageReceived, Messagetype "
    sSQL = sSQL & "FROM Message "
    sSQL = sSQL & "WHERE Message.ClinicalTrialID = -1 "
    sSQL = sSQL & "ORDER BY MessageTimeStamp DESC"
    Set rsMessages = New ADODB.Recordset
    rsMessages.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

    If Not rsMessages.EOF Then
        'return most rescent date
        dDate = rsMessages!MessageTimeStamp
        sDate = LocalNumToStandard(CStr(dDate))
    Else
        'if there are no records then return old date (1/1/1980 00:00)
        sDate = "29221"
    End If
    
    GetLastTransferDate = sDate
    
Exit Function
Errlabel:
    'even though there was an error still return old date so that data transfer can continue (will just cause the reports to be downloaded)
    GetLastTransferDate = "29221"
    sErrMsg = "Error getting last transfer date for reports." & vbCrLf & " Error Description: " & Err.Description & vbCrLf & "Error Number: " & Err.Number
End Function

'---------------------------------------------------------------------
Private Function ReplaceInvalidCharsForDataXfer(sErrMsg As String) As String
'---------------------------------------------------------------------
Dim sMessage As String
    
    sMessage = Replace(sErrMsg, "'", "")
    sMessage = Replace(sMessage, "~", "Tilde")
    sMessage = Replace(sMessage, "|", "Pipe")
    
    ReplaceInvalidCharsForDataXfer = sMessage
    
End Function

'---------------------------------------------------------------------
Private Sub WebBrowser_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
' revisions
' ic 20/11/2006 added check for ie7 format url
'---------------------------------------------------------------------

        'ic 20/11/2006 added check for ie7 format url
        'we are already doing something
        If Left(URL, 15) = "about:blankVBfn" _
        Or Left(URL, 10) = "about:VBfn" Then
            Cancel = True
        End If
        
End Sub

'---------------------------------------------------------------------
Public Function DisplayGetPduMessages() As Boolean
'---------------------------------------------------------------------
' REVISIONS
'---------------------------------------------------------------------
Dim sMessageText As String
' This gets written by the ASP
Const sNO_PDU_MESSAGES = "There are no PDU messages to download"

    Call SetConnectionStatus(csDownloadingPduDataFromServer)
    Call SetTextBoxDisplay("Checking for PDU messages on server " & Me.MACROServerDesc & "...", trStepStarted)
   
    'DoEvents to ensure that button is fully displayed
    DoEvents
    
    On Error GoTo Timeout
    
    'Get pdu messages if there are any
    sMessageText = PostDataHTTP(Me.HTTPAddress & msPduMessageListURL & "?site=" & Me.Site, mnDataTransferTimeout)
    
    On Error GoTo ErrHandler
        
    If sMessageText = "" Then
        GoTo Timeout
    ElseIf Left(sMessageText, 5) = "ERROR" Or InStr(1, LCase(sMessageText), "<html") > 0 Then
        GoTo ASPError
    ElseIf sMessageText = sNO_PDU_MESSAGES Then
        ' no pdu messages to download
        Call SetTextBoxDisplay(sMessageText, trStepCompleted)
    ElseIf sMessageText <> sNO_PDU_MESSAGES Then
        ' download pdu messages
        Call SetConnectionStatus(csDownloadingPduDataFromServer)
        Call SetTextBoxDisplay(sMessageText, trTransferData)

        'if there is a connection error in this routine then exit
        If Not DownloadPduMessages Then GoTo Timeout

    End If
    
    If mbCancelled Then
        'Call SetTextBoxDisplay("Transfer cancelled", trTransferCancelled)
        DisplayGetPduMessages = False
        Exit Function
    End If
    
    DisplayGetPduMessages = True
    
Exit Function

Timeout:
    DisplayTimeoutMessage
    DisplayGetPduMessages = False
Exit Function

ASPError:
    Call SetConnectionStatus(csError)
    'changed error handling to write to data transfer form log
    Call SetTextBoxDisplay("The MACRO server is not responding correctly. Please contact your system administrator.", trTransferError)
    cmdCancel.Caption = "&Close"
    DisplayGetPduMessages = False
    mbConnectionOK = False
    mbTransferOK = False
Exit Function

ErrHandler:
  mbTransferOK = False
  Call SetTextBoxDisplay("An error was encountered while writing PDU messages recieved from the server" & vbCrLf & "Error Description: " & Err.Description & "Error Number: " & Err.Number, trStepCompleted, mbTransferOK)
  DisplayGetPduMessages = True 'set to true as can still continue with the rest of data transfer
End Function

'---------------------------------------------------------------------
Private Function DownloadPduMessages() As Boolean
'---------------------------------------------------------------------
' Download Pdu messages
'---------------------------------------------------------------------
' REVISIONS
' DPH 22/02/2005 - Bug 2534 check for filename and delete if exists
' DPH 18/04/2005 - Changed PDU file download method
'---------------------------------------------------------------------
Dim msNextMessage As String
Dim msMessageId As String
Dim sTrialSite As String
Dim msMessageType As String
Dim msMessageParameters As String
Dim msMessageBody As String
Dim msPDUFile As String
Dim mbPDU() As Byte
Dim mnFreeFile As Integer
Dim sPDUFileDescription As String
Dim asPduFileList() As String
Dim nFileNumber As Integer
Dim i As Integer
Dim sLaunchPduClient As String
Dim vPduMessage As Variant
Dim sErrMsg As String

    On Error GoTo ErrHandler

    cmdCancel.Visible = True
    Call SetConnectionStatus(csDownloadingPduDataFromServer)
    
    nFileNumber = 0
    
    On Error GoTo Timeout
    
    ' DPH 01/05/2002 - Changed call so throws error to be caught gracefully rather than MACRO error
    msNextMessage = PostDataHTTP(Me.HTTPAddress & msGetNextPduMessageURL & "?site=" & Me.Site, 180)
    
    On Error GoTo ErrHandler

    'split the pdu messages from the possible error message return
    vPduMessage = Split(msNextMessage, gsERRMSG_SEPARATOR)
    ' pdu message
    msNextMessage = vPduMessage(0)
    ' error message
    sErrMsg = vPduMessage(1)

    'Collect Error from dll
    If sErrMsg <> "" Then GoTo ErrHandler

    ' DPH 01/05/2002 - Pick up if an error has occurred in the ASP page
    If Left(msNextMessage, 5) = "ERROR" Or InStr(1, LCase(msNextMessage), "<html") > 0 Then
        Call SetTextBoxDisplay("PDU Messages could not be downloaded", trStepCompleted, False)
        ' log error
        gLog gsDOWNLOAD_MESG, "PDU Messages could not be downloaded."
        mbTransferOK = False
        Exit Function
    End If

    Do While msNextMessage <> "."
        msMessageId = ExtractFirstItemFromList(msNextMessage, "<br>")
        msMessageParameters = ExtractFirstItemFromList(msNextMessage, "<br>")
        msMessageBody = ExtractFirstItemFromList(msNextMessage, "<br>")
        DoEvents
        
        On Error GoTo Timeout
        
        'for pdu messages msMessageParameters contains the pdu file that
        'needs to be downloaded
        If RTrim(msMessageParameters) <> "" Then
            'extract the PDU filename from the cab file name
            msPDUFile = msMessageParameters
            ' edit the messagebody parameter to display description
            ' remove first 2 chars
            msMessageBody = Right(msMessageBody, Len(msMessageBody) - 2)
            If InStr(1, msMessageBody, "has been deployed to this site.") > 0 Then
                msMessageBody = Left(msMessageBody, InStr(1, msMessageBody, "has been deployed to this site.") - 1)
            End If
            Call SetTextBoxDisplay("Downloading " & msMessageBody & " from " & Me.MACROServerDesc, trTransferData)
            
            ' DPH 22/02/2005 - Bug 2534 check for filename and delete if exists
            If FolderExistence(gsTEMP_PATH & msPDUFile, True) Then
                Kill gsTEMP_PATH & msPDUFile
            End If
            
            ' new pdu file download - is collected in chunks from server
            ' as may be a large file
            CollectPduMessage msPDUFile, gsTEMP_PATH, sErrMsg
            
            If sErrMsg <> "" Then GoTo ErrHandler

            On Error GoTo ErrHandler
            
            ReDim Preserve asPduFileList(nFileNumber)
            asPduFileList(nFileNumber) = gsTEMP_PATH & msPDUFile
            nFileNumber = nFileNumber + 1
        End If
                        
        
        On Error GoTo Timeout
        
        'Note that msGetNextPdumessageURL is passed the messageId that has just been
        'successfully processed, for the purpose of setting it to read
        msNextMessage = PostDataHTTP(Me.HTTPAddress & msGetNextPduMessageURL & "?site=" & Me.Site & "&previousmessageid=" & msMessageId & "&trialoffice=" & Me.MACROServerDesc, 180)
    
        On Error GoTo ErrHandler
        
        'split the pdu messages from the possible error message return
        vPduMessage = Split(msNextMessage, gsERRMSG_SEPARATOR)
        ' pdu message
        msNextMessage = vPduMessage(0)
        ' error message
        sErrMsg = vPduMessage(1)
    
        'Collect Error from dll
        If sErrMsg <> "" Then GoTo ErrHandler
    
        ' Pick up if an error has occurred in the ASP page
        If Left(msNextMessage, 5) = "ERROR" Or InStr(1, LCase(msNextMessage), "<html") > 0 Then
            Call SetTextBoxDisplay("PDU Messages could not be downloaded", trStepCompleted, False)
            ' log error
            gLog gsDOWNLOAD_MESG, "PDU Messages could not be downloaded."
            mbTransferOK = False
            Exit Function
        End If
    
    Loop
    
    ' put together string to launch pdu client with
    sLaunchPduClient = GetMACROSetting(MACRO_SETTING_PDUEXE_LOCATION, "pdu.exe")
    ' Loop through File List & collect file sizes
    For i = 0 To UBound(asPduFileList)
        sLaunchPduClient = sLaunchPduClient & " """ & asPduFileList(i) & """"
    Next
    ' launch PDU client process - do not wait for response
    If ExecCmdNoWait(sLaunchPduClient) = 0 Then
        Call SetTextBoxDisplay("PDU Client process launch failure.", trTransferData, False)
        mbTransferOK = False
    End If
    
    If Me.Visible = True Then
        
        cmdCancel.Caption = "&Cancel"
        cmdCancel.Visible = True

        Call SetTextBoxDisplay("All PDU messages from " & Me.MACROServerDesc & " downloaded", trStepCompleted)

    End If
    
    Me.MousePointer = vbDefault
    DownloadPduMessages = True
    
Exit Function

Timeout:
    DownloadPduMessages = False
Exit Function

ErrHandler:
    If sErrMsg = "" Then
        sErrMsg = "Error Description: " & Err.Description & ", Error Number: " & Err.Number
    End If
    Call SetTextBoxDisplay("Pdu message download error, " & sErrMsg, trStepCompleted, False)
    Call gLog(gsPDUMSG_DOWNLOAD_ERR, sErrMsg)
    mbTransferOK = False
End Function

'---------------------------------------------------------------------
Private Function CollectPduMessage(sFile As String, sDirectoryToSaveTo As String, _
                                ByRef sErrMsg As String) As Boolean
'---------------------------------------------------------------------
' Download Pdu file from server in small chunks as file may be large
'---------------------------------------------------------------------
Dim sFileToSaveTo As String
Dim byteFile() As Byte
Dim sHTTPAddress As String
Dim sCompleteAddress As String
Dim mnFreeFile As Integer
Dim lFilePos As Long
Dim lMaxChunkSize As Long
Dim bCompleteRead As Boolean
Dim lBytesRead As Long
Dim lFooterStart As Long
Dim sFooter As String
Dim bOK As Boolean
    
    On Error GoTo ErrHandler
    
    bOK = False
    
    sFileToSaveTo = sDirectoryToSaveTo & sFile
    
    ' set up address
    sHTTPAddress = Me.HTTPAddress & msDownloadPduFileURL
    
    ' set up initial fileposition
    lFilePos = 0
    ' set up max chunksize
    lMaxChunkSize = 524288 '(512k)
    
    ' set address
    sCompleteAddress = sHTTPAddress & "?filename=" & sFile & _
                        "&lastfilepos=" & lFilePos & "&maxchunk=" & lMaxChunkSize
    
    On Error GoTo Timeout
    
    ' get first block
    byteFile() = GetDataHTTP(sCompleteAddress, 600)
    
    On Error GoTo ErrHandler
    
    ' complete read initialise
    bCompleteRead = False

    ' if first read ok open file
    mnFreeFile = FreeFile
    Open sFileToSaveTo For Binary Access Write As #mnFreeFile
    While Not bCompleteRead
        ' handle footer
        ' error handling
        lFooterStart = UBound(byteFile) - 2
        If lFooterStart < 0 Then
            ' error
            sFooter = "ERR"
        Else
            ' extract footer
            sFooter = Chr(byteFile(lFooterStart)) & Chr(byteFile(lFooterStart + 1)) & Chr(byteFile(lFooterStart + 2))
        End If
        Select Case sFooter
            Case "SUC"
                ' reduce array to exclude header
                ReDim Preserve byteFile(lFooterStart - 1)
                ' write data to file
                Put #mnFreeFile, , byteFile()
                
                ' if have read less than expected buffer size file is complete
                lBytesRead = UBound(byteFile)
                If LBound(byteFile) = 0 Then
                    lBytesRead = lBytesRead + 1
                End If
                If (lBytesRead < lMaxChunkSize) Then
                    bCompleteRead = True
                    bOK = True
                Else
                    ' read again
                    ' update file position
                    lFilePos = lFilePos + lBytesRead
                    ' set address
                    sCompleteAddress = sHTTPAddress & "?filename=" & sFile & "&lastfilepos=" & lFilePos & "&maxchunk=" & lMaxChunkSize
                    ' get next block
                    On Error GoTo Timeout
    
                    byteFile() = GetDataHTTP(sCompleteAddress, 600)
                    
                    On Error GoTo ErrHandler
                End If
            Case "EMP"
                ' empty so read complete
                bCompleteRead = True
            Case "ERR"
                ' error handle
                sErrMsg = "Error Downloading binary PDU file " & sFile & " from server " & Me.MACROServerDesc
                GoTo ErrHandler
            Case Else
                ' must be an error
                sErrMsg = "Error Downloading binary PDU file " & sFile & " from server " & Me.MACROServerDesc
                GoTo ErrHandler
        End Select
    Wend
    
    ' close file
    Close #mnFreeFile
    
    ' return ok flag
    CollectPduMessage = bOK
Exit Function

Timeout:
    sErrMsg = "Timeout downloading PDU file " & sFile
    CollectPduMessage = False
Exit Function

ErrHandler:
    If sErrMsg = "" Then
        sErrMsg = "" & "Error Description: " & Err.Description & ", Error Number: " & Err.Number
    End If
    CollectPduMessage = False
End Function

'---------------------------------------------------------------------
Public Sub DisplaySendPduMessages()
'---------------------------------------------------------------------
' Upload Pdu messages - wait for response using ExecCmd call
'---------------------------------------------------------------------
' REVISIONS
'---------------------------------------------------------------------
Dim sLaunchPduClient As String

    On Error GoTo ErrHandler
    
    Call SetConnectionStatus(csSendingPduDataToServer)
    Call SetTextBoxDisplay("Sending PDU messages to server...", trStepStarted)
   
    'DoEvents to ensure that button is fully displayed
    DoEvents
    
    ' set PDU Client call
    sLaunchPduClient = GetMACROSetting(MACRO_SETTING_PDUEXE_LOCATION, "pdu.exe") & " upload"
    
    ' execute Pdu client launch
    If (ExecCmdExitCode(sLaunchPduClient) = 0) Then
        Call SetTextBoxDisplay("PDU Client messages sent to server", trStepCompleted)
    Else
        Call SetTextBoxDisplay("PDU Client messages NOT sent to server", trStepCompleted, False)
        mbTransferOK = False
    End If
    
Exit Sub

ErrHandler:
  mbTransferOK = False
  Call SetTextBoxDisplay("An error was encountered whilst launching the PDU process to write PDU messages to the server" & vbCrLf & "Error Description: " & Err.Description & "Error Number: " & Err.Number, trStepCompleted, mbTransferOK)
End Sub
