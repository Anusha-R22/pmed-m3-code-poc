VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmErrors 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "This is set at runtime"
   ClientHeight    =   7575
   ClientLeft      =   3585
   ClientTop       =   4545
   ClientWidth     =   9210
   ControlBox      =   0   'False
   Icon            =   "frmErrors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid grdErrors 
      Height          =   2115
      Left            =   60
      TabIndex        =   14
      Top             =   5400
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   3731
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   3
   End
   Begin VB.TextBox txtUserComment 
      Height          =   1695
      Left            =   5520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   720
      Width           =   3615
   End
   Begin RichTextLib.RichTextBox rtbErrMsg 
      Height          =   3495
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   6165
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmErrors.frx":0442
   End
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   3960
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetTC1 
      Left            =   1200
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   5520
      TabIndex        =   8
      Top             =   3480
      Width           =   3615
      Begin VB.CommandButton cmdOnlineSupport 
         Caption         =   "&Online Support"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print Report"
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   2760
      TabIndex        =   2
      Top             =   4560
      Width           =   6375
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit Application"
         Height          =   375
         Left            =   4800
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdErrDetails 
         Caption         =   "Error &Details"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdTryAgain 
         Caption         =   "Try &Again"
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Please add all additional information about the bug and how it occurred."
      Height          =   495
      Left            =   5520
      TabIndex        =   13
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmErrors.frx":04C4
      Height          =   975
      Left            =   5520
      TabIndex        =   7
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "You can try again or exit from the application."
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   4320
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "The error number and description are:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "An unexpected error has occurred."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   File:       frmErrors.frm
'   Author:     Zulfiqar Ahmed, September 2001
'   Purpose:    Allows the user to decide what to do if the application
'               encounters an unhandled error has functionality to log the error
'               on the Infermed website.
'
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'
'   Revisions:

' TA 09/10/2001: QueryUnload function removed
' NCJ 16 Jan 02 - Ensure print report finishes properly
' TA 17/1/02: VTRACK buglist build 1.0.3 Bug 61 - set the caption to the application name
' TA 18/1/2002 DCBB 2.2.7.7: Decrypt password for online support
' TA 7/1/03 - ignore errors when user not logged in
' RS 10/02/2003: cmdOnlineSupport_Click: Used 'True' to indicate WWWIO dll used, changed to False
'----------------------------------------------------------------------------------------'
Option Explicit
Option Compare Binary
Option Base 0

Private mnTrappedErrNum As Long
Private msTrappedErrDesc As String
Private msTrappedErrForm As String
Private msTrappedErrProc As String
Private msTrappedErrModule As String
Private msTrappedErrorSource As String
Private msObjectName As String
Public gOnErrorAction As OnErrorAction

'---------------------------------------------------------------------
Private Sub cmdErrDetails_Click()
'---------------------------------------------------------------------
'Display or hide the grid
'---------------------------------------------------------------------
    'check the height to see if the grid is already visible
    'if not then change the form height to display grid or vice versa
    
    If Me.Height = 5740 Then
        Me.Height = 7950
        rtbErrMsg.Text = rtbErrMsg.Text & AddErrorsToRTB(Me)
    ElseIf Me.Height = 7950 Then
        Me.Height = 5740
        RefreshForm
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------------------
' Fill the textbox with the error message.
'---------------------------------------------------------------------
Dim iGridWidth As Integer

    cmdOnlineSupport.Enabled = False
    
    'TA 16/01/2003: unload hourglass form if it loaded
    UnloadfrmHourglass
    
    FormCentre Me
    
    Me.Icon = frmMenu.Icon
    Me.BackColor = glFormColour
    Me.Height = 5740
    Screen.MousePointer = vbDefault
    iGridWidth = grdErrors.Width
    grdErrors.ColWidth(0) = 2620
    grdErrors.ColWidth(1) = 3640
    grdErrors.Row = 0
    grdErrors.Col = 0
    grdErrors.Text = "Object Name"
    grdErrors.Col = 1
    grdErrors.Text = "Method Name"
    
    'TA 17/1/02: VTRACK buglist build 1.0.3 Bug 61 - set the caption to the application name
    Me.Caption = GetApplicationTitle & " Error Report"
    cmdErrDetails_Click
End Sub

'------------------------------------------------------------------------------'
Private Sub cmdTryAgain_Click()
'------------------------------------------------------------------------------'
' on an error occuring try to run from the line which caused the error
'------------------------------------------------------------------------------'
    gOnErrorAction = OnErrorAction.Retry
    Me.Hide
End Sub


'------------------------------------------------------------------------------'
Private Sub cmdExit_Click()
'------------------------------------------------------------------------------'
' Exit Macro
'------------------------------------------------------------------------------'
    
    gOnErrorAction = OnErrorAction.QuitMACRO
    Me.Hide

End Sub

'------------------------------------------------------------------------------'
Private Sub cmdPrint_Click()
'------------------------------------------------------------------------------'
' Show the printer facilities for the pc
'------------------------------------------------------------------------------'
Dim sPrevMessage As String

   On Error GoTo DoNothing
    
    'WillC 15/2/00 Append user comments to error message
    RefreshForm
    sPrevMessage = rtbErrMsg.Text
    'add the errors from the grid to the rich text box for printing
    rtbErrMsg.Text = rtbErrMsg.Text & AddErrorsToRTB(Me)
    
    dlg1.Flags = cdlPDReturnDC + cdlPDNoPageNums

    If rtbErrMsg.SelLength = 0 Then
        dlg1.Flags = dlg1.Flags + cdlPDAllPages
    Else
        dlg1.Flags = dlg1.Flags + cdlPDSelection
    End If

    dlg1.CancelError = True
    dlg1.ShowPrinter
    ' Initialise printer
    Printer.Print ""
    rtbErrMsg.SelPrint dlg1.hDc
   
    ' NCJ 16 Jan 02 - Finish the print job
    Printer.EndDoc
    
    'restore the error box back to its original state
    'i.e. remove errors that were added from the grid
    rtbErrMsg.Text = ""
    rtbErrMsg.Text = sPrevMessage

DoNothing:
'resotre if user pressed cancel button
If Err.Number = 32755 Then
    rtbErrMsg.Text = ""
    rtbErrMsg.Text = sPrevMessage
End If
Exit Sub

End Sub

'------------------------------------------------------------------------------'
Private Sub cmdOnlineSupport_Click()
'------------------------------------------------------------------------------'
'You use the replace function to put +'s instead of spaces for the asp string
'components to show them as spaces on the website.Then write to the website problem.asp
'------------------------------------------------------------------------------'

Dim sURL As String
Dim sSupportUser As String
Dim sSupportPassword As String
Dim sErrorNumber As String
Dim sErrorDescription As String
Dim sErrorSource As String
Dim sASPString As String
Dim sRegisteredName As String
Dim sRegisteredOrg As String
Dim sSQL As String
Dim rsSupportDetails As ADODB.Recordset
Dim msResult As String
Dim sTitleOfApp As String
Dim sUserComment As String

    On Error GoTo dbErrors
    
    'WillC 9/5/00
    HourglassOn
    cmdOnlineSupport.Enabled = False

    'WillC 15/2/00 Append user comments to error message
    RefreshForm
    
    'add errors from the error grid to msTrappedErrDesc
    msTrappedErrDesc = msTrappedErrDesc & AddErrorsToRTB(Me)
    
    'WillC 5/5/00 added database type using GetDBType to help with the error messages
    sErrorNumber = Trim(str(mnTrappedErrNum))
    sErrorDescription = Replace(msTrappedErrDesc, " ", "+")
    sUserComment = "Database+Type:+" & GetDBType & "++User+comments:+" & Replace(txtUserComment.Text, " ", "+")
    sErrorSource = Replace(Err.Source, " ", "+")
    ' RS 10/02/2003: Changed WWWDLL parameter from True to False
    sRegisteredName = Replace(GetMACROPCSetting(mpcAuthorisedUser, "Unknown", False), " ", "+")
    ' DPH 10/05/2002 - Need to use licenced company name
    'sRegisteredOrg = "On-line+Support"
    ' RS 10/02/2003: Changed WWWDLL parameter from True to False
    sRegisteredOrg = Replace(GetMACROPCSetting(mpcOrganisation, "Unknown", False), " ", "+")
    
    sSQL = "SELECT * FROM OnlineSupport"
    Set rsSupportDetails = New ADODB.Recordset
    rsSupportDetails.Open sSQL, SecurityADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    With rsSupportDetails
        sURL = rsSupportDetails!SupportURL
        sSupportUser = rsSupportDetails!SupportUserName
        'TA 18/1/2002 DCBB 2.2.7.7: Decrypt password
        sSupportPassword = Crypt(rsSupportDetails!SupportUserPassWord)
    End With
    Set rsSupportDetails = Nothing
     

    sTitleOfApp = GetApplicationTitle
    
    If goUser.UserNameFull <> "" Then
       sRegisteredName = goUser.UserNameFull
       sRegisteredName = Replace(sRegisteredName, " ", "+")
    Else
       sRegisteredName = "Unavailable"
      End If
      
    sTitleOfApp = Replace(sTitleOfApp, " ", "+")

    sASPString = "txtName=" & sRegisteredName & "&CboOrg="
    sASPString = sASPString & sRegisteredOrg & "&GroupType="
    sASPString = sASPString & "Bug&GroupPriority=High&txtDescription=" & sErrorNumber
    sASPString = sASPString & "+" & sErrorDescription & "+" & sErrorSource & "+" & msObjectName & "+" & msTrappedErrProc & "+" & sUserComment & "&CboApplication="
    sASPString = sASPString & sTitleOfApp & "&txtVersionFound=" & App.Major & "." & App.Minor & "." & App.Revision
    sASPString = sASPString & "&checkalpha=0"
   
    On Error GoTo InetTC1Errors

    InetTC1.Protocol = icHTTP
    InetTC1.URL = sURL
    InetTC1.UserName = sSupportUser
    InetTC1.Password = sSupportPassword

    InetTC1.Execute , "POST", sASPString, "Content-Type: application/x-www-form-urlencoded"
    Do Until InetTC1.StillExecuting = False
        DoEvents
    Loop
    
    'WillC 9/5/00 If the submission has been succesful we get something in msResult
    msResult = InetTC1.GetChunk(102400)
    'WillC 9/5/00 If the submission has been succesful send a confirmation
    If msResult <> "" Then
        DialogInformation "Your error has been submitted to the InferMed support site"
    Else        'if not tell the user to make sure they have a valid connection.
        DialogError " Unable to submit this error to the InferMed support site." & vbCrLf _
            & "  Please ensure your machine has a valid connection to the internet."

    End If
   
   HourglassOff
Exit Sub
dbErrors:
       HourglassOff
       DialogInformation " There has been an error connecting to the database." & vbCrLf _
             & " Please check the database path"
       
Exit Sub
'   WillC 3/2/00 Added error handler to ensure the user has a connection to the Web.
InetTC1Errors:
        HourglassOff
        Select Case Err.Number
            Case 35764  ' Still executing last request
                DialogInformation " The InferMed support site may be busy at the moment." & vbCrLf _
                    & "  We are still executing the last request."
                    cmdOnlineSupport.Enabled = False
            Case Else
                DialogWarning " Please ensure your machine has a connection to the internet."
                Exit Sub
        End Select
                                          
End Sub

'------------------------------------------------------------------------------'
Public Sub ProcessErrors(sObjectName As String, nTrappedErrNum As Long, sTrappedErrDesc As String, sProcName As String, sSource As String)
'------------------------------------------------------------------------------'
' Get the error details and pass them to be shown on the form.
' This new procedure is based on FormRefreshMe but instead of passing the form
' object as a parameter, we will pass the object name
'------------------------------------------------------------------------------'
Dim sLocation As String
Dim sProc As String
Dim iPointLocation As Integer

On Error GoTo ErrLabel


    Screen.MousePointer = vbDefault

    If nTrappedErrNum > vbObjectError And nTrappedErrNum < vbObjectError + 65536 Then
        mnTrappedErrNum = nTrappedErrNum - vbObjectError
    Else
        mnTrappedErrNum = nTrappedErrNum
    End If
    
    If Len(sSource) > 0 Then
        iPointLocation = InStr(1, sSource, ".")
        If iPointLocation > 0 Then
            sLocation = Left$(sSource, iPointLocation - 1)
            sProc = Right$(sSource, Len(sSource) - iPointLocation)
            msTrappedErrModule = sLocation
            msTrappedErrDesc = sTrappedErrDesc
            msTrappedErrProc = sProc
        Else
            msTrappedErrDesc = sTrappedErrDesc
            msObjectName = sObjectName
            msTrappedErrProc = sProcName
        End If
    Else
    End If
    
    RefreshForm
    
Exit Sub
    
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmErrors.ProcessErrors"
    
End Sub


'------------------------------------------------------------------------------'
Private Sub RefreshForm()
'------------------------------------------------------------------------------'
' Show the error details in the text box depending on whether hte error came from
' a form or a module.
'------------------------------------------------------------------------------'
Dim sMsgText As String
Dim sProcedureName As String
Dim sTitleOfApp As String
Dim sFullMsg As String
Dim sHeader As String
Dim sComments As String
Dim sDatabaseType As String



    
    rtbErrMsg.Text = ""
       
    sTitleOfApp = GetApplicationTitle
    

    If txtUserComment.Text <> vbNullString Then
       sComments = "User comments: " & txtUserComment.Text
    End If

    sMsgText = "Error Number: " & str(mnTrappedErrNum) & vbCrLf _
             & "Error Description: " & msTrappedErrDesc & vbCrLf _
             & "Error Source: " & sTitleOfApp & vbCrLf _
             & "Error Location: " & msObjectName & vbCrLf _
             & "In Routine: " & msTrappedErrProc & vbCrLf _
             & vbCrLf _
             & sComments
       
    'TA 17/1/02: part of VTRACK buglist build 1.0.3 Bug 6
    sHeader = sTitleOfApp & " ERROR REPORT "
      
    sFullMsg = sHeader & vbCrLf & vbCrLf
    sFullMsg = sFullMsg & "DATE: " & Format(Now, "dd/mm/yyyy") & vbCrLf
    sFullMsg = sFullMsg & "TIME: " & Time & vbCrLf
    
    
    'TA: next two lines will cause error if not logged in
    On Error Resume Next
    sFullMsg = sFullMsg & "USER: " & goUser.UserName & vbCrLf
    sFullMsg = sFullMsg & "DATABASE: " & GetDBType & vbCrLf
    On Error GoTo 0
    
    sFullMsg = sFullMsg & "VERSION: " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf
      
    rtbErrMsg.Text = sFullMsg & vbCrLf & sMsgText
    
End Sub

'------------------------------------------------------------------------------'
Private Sub txtUserComment_Change()
'------------------------------------------------------------------------------'
'Ensure the user enters a comment before creating an online support instance.
'------------------------------------------------------------------------------'

    If txtUserComment.Text <> vbNullString Then
        cmdOnlineSupport.Enabled = True
    Else
        cmdOnlineSupport.Enabled = False
    End If
    
        
End Sub

'------------------------------------------------------------------------------'
Private Function GetDBType() As String
'------------------------------------------------------------------------------'
' Databasetype for error messages
'------------------------------------------------------------------------------'
Dim sDatabaseType As String

    Select Case goUser.Database.DatabaseType
        Case MACRODatabaseType.Access
            sDatabaseType = "Access"
        Case MACRODatabaseType.sqlserver
            sDatabaseType = "SQLServer"
        Case MACRODatabaseType.SQLServer70
            sDatabaseType = "SQLServer70"
        Case MACRODatabaseType.Oracle80
            sDatabaseType = "Oracle"
    End Select

    GetDBType = sDatabaseType
    
End Function

