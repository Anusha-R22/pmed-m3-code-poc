VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmTimeSynch 
   Caption         =   "Remote Time Synchronisation"
   ClientHeight    =   1860
   ClientLeft      =   6075
   ClientTop       =   5400
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   5205
   Begin VB.Frame Frame2 
      Height          =   2205
      Left            =   60
      TabIndex        =   5
      Top             =   1860
      Width           =   7650
      Begin VB.Label lblServerResult 
         Caption         =   "lblServerResult"
         Height          =   600
         Left            =   645
         TabIndex        =   13
         Top             =   1470
         Width           =   6885
      End
      Begin VB.Label Label3 
         Caption         =   "Time synchronisation server results :"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   1185
         Width           =   2835
      End
      Begin VB.Label lblLocalTZ 
         Caption         =   "lbllocaltz"
         Height          =   255
         Left            =   3060
         TabIndex        =   11
         Top             =   240
         Width           =   4470
      End
      Begin VB.Label Label2 
         Caption         =   "Local computer time zone"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2235
      End
      Begin VB.Label lblServerTZ 
         Caption         =   "lblServerTZ"
         Height          =   255
         Left            =   3060
         TabIndex        =   9
         Top             =   840
         Width           =   4470
      End
      Begin VB.Label lblAddress 
         Caption         =   "lblAddress"
         Height          =   255
         Left            =   3060
         TabIndex        =   8
         Top             =   540
         Width           =   4470
      End
      Begin VB.Label Label4 
         Caption         =   "Time synchronisation server address"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   540
         Width           =   2715
      End
      Begin VB.Label Label3 
         Caption         =   "Time synchronisation server time zone"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   2835
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3900
      TabIndex        =   1
      Top             =   1380
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Wsk1 
      Left            =   240
      Top             =   1380
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   900
      Top             =   1380
   End
   Begin VB.Frame Frame1 
      Caption         =   "Time"
      Height          =   1215
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5055
      Begin VB.CommandButton cmdConnect 
         Caption         =   "&Synchronise"
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   300
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Your system clock will be reset to your correct local time."
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   4575
      End
      Begin VB.Label lblOutput 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmTimeSynch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999-2000, All Rights Reserved
'   File:       frmTimeSynchroniser.frm
'   Author:     Will Casey, October 1999
'   Purpose:    To enable a users machine to synchronise with a trusted external time
'               server...
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'  WillC 15/6/00 Added the  api call to find which zone the machine is set up for and then do the
'   maths to add or subtract the necessary amount of time in minutes to get the correct local time
'   by adding the difference to the time returned by the navy clock ie
'   if it is 12 noon GMT and the machine is set for Teheran + 3.5 hours then the clock is set to 3:30
'   local time taking into consideration Daylight saving time
'   NCJ 12/9/00 - Reset form icon because of VB errors
'   NCJ 27/10/00 - Removed Help button (because Help isnow in DM)
'   NCJ 9 Nov 00 - Fixed several bugs (time zone and daylight saving now correctly dealt with)
'   DPH 29/05/2002 - Altered Time Host Setting to tyco from tock
'   ZA 17/06/2002 - fixed CBB 2.2.16.17 - no permission to change clock
'
'------------------------------------------------------------------------------------'
' NOTE: Much of this code was copied from the VB Time Zone Example to be found at:
'   \\IMED1\Dev\TimeZoneExample
'------------------------------------------------------------------------------------'

Option Explicit

Private Declare Function GetTimeZoneInformation _
   Lib "kernel32" (lpTimeZoneInformation As _
   TIME_ZONE_INFORMATION) As Long

Private Type SYSTEMTIME
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
   Bias As Long
   StandardName(0 To 63) As Byte
   StandardDate As SYSTEMTIME
   StandardBias As Long
   DaylightName(0 To 63) As Byte
   DaylightDate As SYSTEMTIME
   DaylightBias As Long
End Type

Private Const TIME_ZONE_ID_INVALID = &HFFFFFFFF
Private Const TIME_ZONE_ID_UNKNOWN = 0
Private Const TIME_ZONE_ID_STANDARD = 1
Private Const TIME_ZONE_ID_DAYLIGHT = 2

Private mlDifferenceInMins As Long
Private mlDaylightSavingInMins As Long

'rts server
Private msRemoteTimeSyncServer As String
'rts server offset
Private mlRemoteTimeOffSet As Long

'-----------------------------------------------------------------------------------
Private Sub ServerIsBusy()
'-----------------------------------------------------------------------------------
' Tell the user that the time server is busy
'-----------------------------------------------------------------------------------
            
    Call DialogInformation("The remote time server is busy. Please try again later.")
    lblServerResult.Caption = "Remote time server was busy"

End Sub

'-----------------------------------------------------------------------------------
Private Sub cmdConnect_Click()
'-----------------------------------------------------------------------------------
' Open the connection to the time server
'-----------------------------------------------------------------------------------

    On Error GoTo ErrHandler
    
        Screen.MousePointer = vbHourglass
        lblOutput.Caption = ""
    Timer1.Enabled = True
    Wsk1.Close
'    CheckTimer
    Wsk1.Connect
    
    cmdConnect.Enabled = False
    
Exit Sub
ErrHandler:
    Screen.MousePointer = vbNormal
    Select Case Err.Number
        Case 40020
            Call ServerIsBusy
    End Select
 
End Sub

' NCJ 27/10/00 - Help button removed
''-----------------------------------------------------------------------------------
'Private Sub cmdHelp_Click()
''-----------------------------------------------------------------------------------
'
''-----------------------------------------------------------------------------------
'  Dim sDocPath As String
'
'    sDocPath = App.Path & "\help\ts\contents.htm"
'    Call ShowDocument(Me.hWnd, sDocPath)
'
'End Sub

'-----------------------------------------------------------------------------------
Private Sub cmdQuit_Click()
'-----------------------------------------------------------------------------------
' Unload
'-----------------------------------------------------------------------------------

    Unload Me
    
End Sub

'-----------------------------------------------------------------------------------
Private Sub Form_Load()
'-----------------------------------------------------------------------------------
' Set the protocol to TCP, the server we are hitting is the US Navy Observatory time
' server this is the same server used by the WINNT TimeServ utility.The port is necessary
' to make the connection.
'-----------------------------------------------------------------------------------
' REVISIONS
' DPH 29/05/2002 - Altered Time Host Setting to tyco from tock & reads from registry
'-----------------------------------------------------------------------------------

    FormCentre Me, frmMenu
    
    Me.Icon = frmMenu.Icon
    Timer1.Interval = 5000 '5 secs
    Timer1.Enabled = False
    
    Wsk1.Protocol = sckTCPProtocol
    '  DPH 29/05/2002 - Altered Time Host Setting to tyco from tock
    '  Will attempt to read from registry firstly before defaulting
    mlRemoteTimeOffSet = CLng(GetMACROSetting(MACRO_SETTING_REMOTETIMESYNCOFFSET, "0"))  ' 7 is MST
    msRemoteTimeSyncServer = GetMACROSetting(MACRO_SETTING_REMOTETIMESYNCSERVER, "time-b.nist.gov")
    
    Wsk1.RemoteHost = msRemoteTimeSyncServer
    Wsk1.RemotePort = "13"
    
    ' NCJ 9/11/00 - Originally in Form_Paint (!!)
    Call ReadTimeZone
    
    
    lblServerResult.Caption = ""
    lblAddress.Caption = msRemoteTimeSyncServer
    lblServerTZ.Caption = "GMT " & IIf(mlRemoteTimeOffSet > 0, "+ ", "- ") & Abs(mlRemoteTimeOffSet)
    lblLocalTZ.Caption = "GMT " & IIf(mlDaylightSavingInMins + mlDifferenceInMins > 0, "- ", "+ ") & Abs((mlDaylightSavingInMins + mlDifferenceInMins) \ 60)
    
End Sub

'-----------------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'-----------------------------------------------------------------------------------
' Close the winsock control down.
'-----------------------------------------------------------------------------------
    
    Wsk1.Close
        Screen.MousePointer = vbNormal
    
End Sub

Private Sub lblAddress_Click()

On Error GoTo ErrLabel
        Wsk1.Close
    If frmInputBox.Display("Remote time server", "Please select (list available at http://www.ativapro.com/timeservers.htm):", msRemoteTimeSyncServer, False) Then
        msRemoteTimeSyncServer = Trim(msRemoteTimeSyncServer)
        Wsk1.RemoteHost = msRemoteTimeSyncServer
        Wsk1.RemotePort = "13"
        Call SetMACROSetting(MACRO_SETTING_REMOTETIMESYNCSERVER, msRemoteTimeSyncServer)
        lblAddress.Caption = msRemoteTimeSyncServer
    End If
    Exit Sub
ErrLabel:
    lblAddress.Caption = msRemoteTimeSyncServer
End Sub

Private Sub lblServerTZ_Click()
    '
Dim sRemoteTimeOffset As String

On Error GoTo ErrLabel
    If frmInputBox.Display("Remote time server time zone", "Please enter timezone offset to GMT in hours", sRemoteTimeOffset, False) Then
        mlRemoteTimeOffSet = CLng(sRemoteTimeOffset)
        lblServerTZ.Caption = "GMT " & IIf(mlRemoteTimeOffSet > 0, "+ ", "- ") & Abs(mlRemoteTimeOffSet)
        Call SetMACROSetting(MACRO_SETTING_REMOTETIMESYNCOFFSET, sRemoteTimeOffset)
    End If
    Exit Sub
ErrLabel:
    lblServerTZ.Caption = "GMT " & IIf(mlRemoteTimeOffSet > 0, "+ ", "- ") & Abs(mlRemoteTimeOffSet)
End Sub

'-----------------------------------------------------------------------------------
Private Sub Timer1_Timer()
'-----------------------------------------------------------------------------------
' If we dont get a response within 10 seconds then pop up the message and ask the
' user to try again later.
'-----------------------------------------------------------------------------------

    On Error GoTo ErrHandler
        Screen.MousePointer = vbNormal
    If lblOutput.Caption = "" Then
        Call ServerIsBusy
        Wsk1.Close
    End If
    Timer1.Enabled = False
    cmdConnect.Enabled = True

Exit Sub
ErrHandler:
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "Timer1_Timer")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'-----------------------------------------------------------------------------------
Private Sub wsk1_DataArrival(ByVal bytesTotal As Long)
'-----------------------------------------------------------------------------------
' When the server sends the date time information back the format is e.g.
' "Thu Nov 18 11:15:05 1999 vbcrlf vbcrlf"
' strip out the time '11:15:05' as a string
' and use this to reset the machine's time.
' NCJ 9 Nov 00 - NOTE We do NOT take the date into consideration
'-----------------------------------------------------------------------------------
Dim sData As String
Dim sTimeString As Variant
Dim lLocalAdjustment As Long
Dim lColonPos As Integer

    On Error GoTo ErrHandler
    
    Screen.MousePointer = vbArrowHourglass
    
    Wsk1.GetData sData
    
    DoEvents
    
    
    lblServerResult.Caption = sData
    ' Get time as hh:mm:ss
    lColonPos = InStr(1, sData, ":")
    ' Back 2 chars for hh
    sTimeString = Mid(sData, lColonPos - 2, Len("hh:mm:ss"))
    
    'Add all the time differences here
    lLocalAdjustment = mlDifferenceInMins + mlDaylightSavingInMins
    'TA 28/03/2004: remove 7 hours as now returned in MST
    lLocalAdjustment = lLocalAdjustment + (mlRemoteTimeOffSet * 60)
    lLocalAdjustment = -lLocalAdjustment
    
    ' Add the difference in minutes to the GMT time returned by ping
    ' Convert it to a string for the display
    ' (Note that this ignores the date so the adjustment will never change the day)
    sTimeString = Format(DateAdd("n", lLocalAdjustment, CDate(sTimeString)), "hh:mm:ss")
    
    lblOutput.Caption = sTimeString
    'Change the clock to the correct local time...
    Time = sTimeString

    Wsk1.Close

Exit Sub

ErrHandler:
    
    'ZA 17/06/2002 CBB 2.2.16.7
    'Throw an error if user doesn't have proper privilege to change system time
    If Err.Number = 70 Then
        Call DialogError("You don't have permission to change the system clock.")
        Exit Sub
    End If
    
  Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "wsk1_DataArrival")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

 
End Sub

'-----------------------------------------------------------------------------------
Private Sub ReadTimeZone()
'-----------------------------------------------------------------------------------
' Get the timezone the computer has been set in
' Find out if Daylight saving time applies
' Set module-level variables
' NCJ 9/11/00 - Changed so that it correctly takes account of Daylight Saving
'-----------------------------------------------------------------------------------
Dim lRet As Long
Dim tz As TIME_ZONE_INFORMATION

    ' Initialise variables
    mlDaylightSavingInMins = 0
    mlDifferenceInMins = 0
    
   ' lRet tells you whether it's Standard time or Daylight Savings time
    lRet = GetTimeZoneInformation(tz)
    
    Select Case lRet
    Case TIME_ZONE_ID_INVALID, TIME_ZONE_ID_UNKNOWN
        ' The call failed for some reason
    Case TIME_ZONE_ID_STANDARD
        ' Standard Time - pick up general time zone difference
        mlDifferenceInMins = tz.Bias
    Case TIME_ZONE_ID_DAYLIGHT
        ' Daylight Savings Time - pick up general time zone difference AND daylight saving
        mlDifferenceInMins = tz.Bias
        mlDaylightSavingInMins = tz.DaylightBias
    End Select

End Sub



