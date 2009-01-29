VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenu 
   Caption         =   "MACRO LISA Test"
   ClientHeight    =   8790
   ClientLeft      =   3045
   ClientTop       =   4020
   ClientWidth     =   12495
   Icon            =   "frmMenuLISATest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   12495
   Begin VB.Frame Frame1 
      Caption         =   "Soapiness"
      Height          =   915
      Left            =   10980
      TabIndex        =   19
      Top             =   1920
      Width           =   1395
      Begin VB.OptionButton optSoap 
         Caption         =   "Use SOAP"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1155
      End
      Begin VB.OptionButton optDLL 
         Caption         =   "Use DLL"
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   1035
      End
   End
   Begin VB.TextBox txtXMLSubject 
      Height          =   315
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   120
      Width           =   7395
   End
   Begin VB.CommandButton cmdQuestion 
      Caption         =   "Question"
      Height          =   315
      Left            =   10980
      TabIndex        =   15
      Top             =   1260
      Width           =   1095
   End
   Begin VB.CommandButton cmdEform 
      Caption         =   "Eform"
      Height          =   315
      Left            =   10980
      TabIndex        =   14
      Top             =   900
      Width           =   1095
   End
   Begin VB.CommandButton cmdVisit 
      Caption         =   "Visit"
      Height          =   315
      Left            =   10980
      TabIndex        =   13
      Top             =   540
      Width           =   1095
   End
   Begin VB.TextBox txtXML 
      Height          =   2655
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   12
      Text            =   "frmMenuLISATest.frx":08CA
      Top             =   480
      Width           =   7395
   End
   Begin VB.Frame fraRevalidation 
      Caption         =   "Subject XML"
      Height          =   4395
      Left            =   60
      TabIndex        =   8
      Top             =   3300
      Width           =   12015
      Begin VB.CommandButton cmdUnlock 
         Caption         =   "Unlock subject"
         Height          =   315
         Left            =   4800
         TabIndex        =   22
         Top             =   300
         Width           =   1815
      End
      Begin VB.CommandButton cmdXMLInput 
         Caption         =   "Do XML Input"
         Height          =   345
         Left            =   180
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdDoXML 
         Caption         =   "Do XML request"
         Height          =   345
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   1305
      End
      Begin VB.TextBox txtMsg 
         Height          =   3495
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Text            =   "frmMenuLISATest.frx":0B4D
         Top             =   720
         Width           =   11775
      End
   End
   Begin VB.Frame fraSubj 
      Caption         =   "Subject"
      Height          =   675
      Left            =   60
      TabIndex        =   6
      Top             =   1575
      Width           =   3195
      Begin VB.TextBox txtSubject 
         Height          =   315
         Left            =   120
         MaxLength       =   255
         TabIndex        =   7
         Text            =   "RED 222"
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.Frame fraStudy 
      Caption         =   "Study"
      Height          =   675
      Left            =   60
      TabIndex        =   4
      Top             =   120
      Width           =   3210
      Begin VB.ComboBox cboStudy 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.Frame fraSite 
      Caption         =   "Site"
      Height          =   675
      Left            =   60
      TabIndex        =   2
      Top             =   855
      Width           =   3195
      Begin VB.ComboBox cboSite 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   345
      Left            =   10860
      TabIndex        =   0
      Top             =   7860
      Width           =   1125
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6180
      Top             =   5100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrSystemIdleTimeout 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   5700
      Top             =   5160
   End
   Begin MSComctlLib.StatusBar sbrMenu 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   8415
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   900
            MinWidth        =   882
            Key             =   "UserKey"
            Object.ToolTipText     =   "Name of current user"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   900
            MinWidth        =   882
            Key             =   "RoleKey"
            Object.ToolTipText     =   "Role of current user"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   900
            MinWidth        =   882
            Key             =   "UserDatabase"
            Object.ToolTipText     =   "Name of current Database"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblFile 
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   7800
      Width           =   9075
   End
   Begin VB.Label lblSubject 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Study/Site/Label"
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   2340
      Width           =   3135
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuFHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHUserGuide 
         Caption         =   "&User Guide"
      End
      Begin VB.Menu mnuSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHAboutMacro 
         Caption         =   "&About MACRO"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
' File:         frmMenuLISATest.frm
' Copyright:    InferMed Ltd. 2003-2007. All Rights Reserved
' Author:       Nicky Johns, August 2003
' Purpose:      Tests the MACRO/LISA interface
'----------------------------------------------------------------------------------------'
' Revisions:
'   NCJ 13-14 Aug 03 - Initial development
'   NCJ 22 Jan 07 - Updated for MACRO 3.0.76
'----------------------------------------------------------------------------------------'

Option Explicit

Public gsSOAP As String

'Private Const msSOAP_ADDRESS = "http://NICKY/LISAII/LISAII.WSDL"
' Nicky's machine for testing
' Private Const msSOAP_ADDRESS = "http://JOHNSN/LISAII/LISAII.WSDL"
' 23 Jan 07 - Try David's machine
 Private Const msSOAP_ADDRESS = "http://HOOKD/LISAII/LISAII.WSDL"

' R3005 = Sub 6, R3008 = Subj 23, R3021 = Subj 45, R3022 = Subj 47,
' R3024 = Subj 49, R3001 = Subj 9, R3047 = Subj 68
'Private Const msTEST_SUBJ_LABEL = "R3008,  HH,  01.01.1985"
'Private Const msTEST_SUBJ_LABEL = "R3021,  CC,  15.02.1990"
Private Const msTEST_SUBJ_LABEL = "R3022,  AA,  03.03.1995"
'Private Const msTEST_SUBJ_LABEL = "R3024,  HC,  01.06.1995"
'Private Const msTEST_SUBJ_LABEL = "R3001,  QQ,  03.03.1993"
'Private Const msTEST_SUBJ_LABEL = "R3047,  KK,  04.04.1989"
'Private Const msTEST_SUBJ_LABEL = "R3005,  TC,  09.04.1995"
'Private Const msTEST_SUBJ_LABEL = "R3007,  HH,  01.01.1995"

Private mlSelStudyId As Long
Private msSelStudyName As String
Private msSelSite As String
Private msSubjLabel As String

' The subject list array
Private mvSubjects As Variant

Private mlMinScaleHeight As Long
Private mlMinScaleWidth As Long

Private mcolWritableSites As Collection

Private moUser As MACROUser

Private Const msALL_SITES = "All Sites"

' The minimum top coord for the Revalidation frame
Private Const mlREVALIDATION_TOP As Long = 3000
' The gap between controls
Private Const mlGAP As Long = 60

' The log file (where appropriate)
Private msLogFile As String

' Our stored lock tokens
Private msLockTokens As String

'--------------------------------------------------------------------
Private Sub cmdExit_Click()
'--------------------------------------------------------------------

    Call mnuFExit_Click

End Sub

'--------------------------------------------------------------------
Public Sub InitialiseMe()
'--------------------------------------------------------------------
' This gets called from Main when MACRO starts up
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    'The following Doevents prevents command buttons ghosting during form load
    DoEvents
    
    If LoadStudies Then
        Call LoadSites
    End If
'    txtSubject.Text = ""

Exit Sub
ErrHandler:
'    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "InitialiseMe", Err.Source) = Retry Then
'        Resume
'    End If
End Sub

'--------------------------------------------------------------------
Public Sub CheckUserRights()
'--------------------------------------------------------------------
' Dummy routine which gets called during MACRO initialisation
'--------------------------------------------------------------------

End Sub

'--------------------------------------------------------------------
Private Function GetLogFileName() As String
'--------------------------------------------------------------------
' Get name of log file for results of revalidation
'--------------------------------------------------------------------
    
    GetLogFileName = App.Path & "\RVLog " & Format(Now, "d-mmm-yy hh-mm-ss") & ".txt"

End Function


'--------------------------------------------------------------------
Private Sub DoTheXMLRequest()
'--------------------------------------------------------------------
' Do the XML thing
'--------------------------------------------------------------------
Dim sLogFile As String
Dim sMsg As String
Dim sRetData As String
Dim oLISA As MACROLISA
Dim oSoapClient As SoapClient30
Dim sglT1 As Single
Dim sglT2 As Single

'Const sSTUDY_NAME = "Demostudy30"
Const sSTUDY_NAME = "ALLR3_UAT_2"

    On Error GoTo Errlabel

    If txtXMLSubject.Text > "" Then
        ' Check they want to do it
        sMsg = "Process this XML data request?"
        If DialogQuestion(sMsg) = vbYes Then
    
            Screen.MousePointer = vbHourglass
            txtMsg.Text = Now & " - Retrieving data, please wait..."
            DoEvents
            
            sglT1 = Timer
            If optDLL.Value = True Then
                ' Use DLL
                Set oLISA = New MACROLISA
                msLockTokens = ""
                txtMsg.Text = "Data retrieval result code = " & _
                                oLISA.GetLISASubjectData(moUser.GetStateHex(False), _
                                    msSelStudyName, msSubjLabel, sRetData, msLockTokens)
                Set oLISA = Nothing
            Else
                ' Use SOAP
                Set oSoapClient = New SoapClient30
                Call oSoapClient.MSSoapInit(gsSOAP)
                oSoapClient.ConnectorProperty("Timeout") = 900000
                txtMsg.Text = "Data retrieval result code = " & _
                                oSoapClient.GetLISASubjectData(moUser.GetStateHex(False), _
                                    msSelStudyName, msSubjLabel, sRetData, msLockTokens)
                Set oSoapClient = Nothing
            End If
            sglT2 = Timer
            Screen.MousePointer = vbNormal
            
            txtMsg.Text = txtMsg.Text & vbCrLf & "Time taken = " & sglT2 - sglT1
            If DialogQuestion("Save returned data to file?") = vbYes Then
                msLogFile = SelectLogFile
                lblFile.Caption = "Output file: " & msLogFile
                If msLogFile > "" Then
                    Call LogToFile(sRetData)
                End If
            Else
                If DialogQuestion("Display returned data?") = vbYes Then
                    txtMsg.Text = txtMsg.Text & vbCrLf & sRetData
                End If
            End If
            If msLockTokens > "" Then
                If DialogQuestion("Display lock tokens?") = vbYes Then
                    txtMsg.Text = txtMsg.Text & vbCrLf & "LOCKS = " & msLockTokens
                End If
            End If

        End If
    Else
        MsgBox "No subject specified!"
    End If
    
Exit Sub
Errlabel:
    txtMsg.Text = "An error occurred in DoTheXMLRequest!" & vbCrLf _
                    & Err.Number & ", " & Err.Description
    Screen.MousePointer = vbNormal

End Sub

'--------------------------------------------------------------------
Private Sub cmdDoXML_Click()
'--------------------------------------------------------------------
' Revalidate subjects
'--------------------------------------------------------------------

    Call DoTheXMLRequest

End Sub

'--------------------------------------------------------------------
Private Sub cmdUnlock_Click()
'--------------------------------------------------------------------
' Unlock the current subject
'--------------------------------------------------------------------
Dim oLISA As MACROLISA
Dim oSoapClient As SoapClient30
Dim sErrMsg As String
Dim sglT1 As Single
Dim sglT2 As Single
Dim sUser As String
Dim sLocks As String

    If msLockTokens = "" Then
        MsgBox "We have no lock tokens!"
    ElseIf MsgBox("Unlock all eforms for this subject?", vbYesNo) = vbYes Then
    
        Screen.MousePointer = vbHourglass
        sLocks = msLockTokens
        sUser = moUser.GetStateHex(False)
        sglT1 = Timer
        
        If optDLL.Value = True Then
            ' Use DLL
            Set oLISA = New MACROLISA
            txtMsg.Text = "Unlocking - Result code = " & oLISA.UnlockSubject(sUser, sLocks, sErrMsg)
            Set oLISA = Nothing
            If sErrMsg > "" Then
                txtMsg.Text = txtMsg.Text & vbCrLf & sErrMsg
            End If
        Else
            ' Use SOAP
            Set oSoapClient = New SoapClient30
            Call oSoapClient.MSSoapInit(gsSOAP)
            txtMsg.Text = "Unlocking - Result code = " & oSoapClient.UnlockSubject(sUser, sLocks, sErrMsg)
            Set oSoapClient = Nothing
            If sErrMsg > "" Then
                txtMsg.Text = txtMsg.Text & vbCrLf & sErrMsg
            End If
        End If
        
        sglT2 = Timer
        Screen.MousePointer = vbNormal
        txtMsg.Text = txtMsg.Text & vbCrLf & "Time taken = " & sglT2 - sglT1
        
        msLockTokens = ""
    End If
    
    
End Sub

'--------------------------------------------------------------------
Private Sub cmdXMLInput_Click()
'--------------------------------------------------------------------
' Send input data
'--------------------------------------------------------------------
Dim sLogFile As String
Dim sMsg As String
Dim sRetData As String
Dim oLISA As MACROLISA
Dim oSoapClient As SoapClient30
Dim sglT1 As Single
Dim sglT2 As Single
Dim sUser As String
Dim sLocks As String

    On Error GoTo Errlabel

    If txtXMLSubject.Text > "" Then
        ' Check they want to do it
        sMsg = "Process this XML data input"
        If msLockTokens = "" Then
            sMsg = sMsg & " (NO lock tokens)?"
        Else
            sMsg = sMsg & " (WITH lock tokens)?"
        End If
        If DialogQuestion(sMsg) = vbYes Then
        
            Screen.MousePointer = vbHourglass
            txtMsg.Text = Now & " - Doing data input, please wait..."
            DoEvents
            
            sUser = moUser.GetStateHex(False)
            sLocks = msLockTokens
            sglT1 = Timer
            If optDLL.Value = True Then
                ' Use DLL
                Set oLISA = New MACROLISA
                txtMsg.Text = "Finished Data Input" & vbCrLf _
                                & "Result code = " _
                                & oLISA.InputLISASubjectData(sUser, _
                                            GetTheXMLRequest, sLocks, sRetData)
                Set oLISA = Nothing
                
            Else
                ' Use SOAP
                Set oSoapClient = New SoapClient30
                Call oSoapClient.MSSoapInit(gsSOAP)
                oSoapClient.ConnectorProperty("Timeout") = 900000
                txtMsg.Text = "Finished Data Input" & vbCrLf _
                                & "Result code = " _
                                & oSoapClient.InputLISASubjectData(sUser, _
                                    GetTheXMLRequest, sLocks, sRetData)
                Set oSoapClient = Nothing
            End If
            sglT2 = Timer
            Screen.MousePointer = vbNormal
            
            txtMsg.Text = txtMsg.Text & vbCrLf & "Time taken = " & sglT2 - sglT1
            txtMsg.Text = txtMsg.Text & vbCrLf & sRetData
            
        End If
    Else
        MsgBox "No subject specified!"
    End If
    
Exit Sub
Errlabel:
    txtMsg.Text = "An error occurred in cmdXMLInput!" & vbCrLf _
                    & Err.Number & ", " & Err.Description
    Screen.MousePointer = vbNormal

End Sub

'--------------------------------------------------------------------
Private Sub Form_Load()
'--------------------------------------------------------------------
Dim sMsg As String
Dim sUserFull As String
Dim sSerialUser As String
Dim lLoginResult  As Long
Dim bSoapy As Boolean

    FormCentre Me
    
    txtMsg.Text = ""
'    txtXML.Text = ""
    
    cmdDoXML.Enabled = False
    cmdXMLInput.Enabled = False
    
    optSoap.Value = False
    optDLL.Value = True
    
    ' Remember our "minimum" size (this is how we start off)
    mlMinScaleHeight = Me.ScaleHeight
    mlMinScaleWidth = Me.ScaleWidth
    
    gsSOAP = msSOAP_ADDRESS
    
    ' Do LISA login
    lLoginResult = frmLISALogin.Display(sMsg, sUserFull, sSerialUser, bSoapy)
    
    If lLoginResult = -1 Then
        ' They cancelled - get out now
        Call mnuFExit_Click
    Else
        txtMsg.Text = "Login result = " & lLoginResult & vbCrLf
        optSoap.Value = bSoapy
        optDLL.Value = Not bSoapy
        If lLoginResult = 0 And sSerialUser > "" Then
            txtMsg.Text = txtMsg.Text & "User = " & sUserFull & vbCrLf
            Set moUser = New MACROUser
            Call moUser.SetStateHex(sSerialUser)
            Call InitialiseMe
        Else
            txtMsg.Text = txtMsg.Text & "Message = " & sMsg & vbCrLf
        End If
    
    '    txtSubject.Text = "RED 222"
        txtSubject.Text = msTEST_SUBJ_LABEL
        Call txtSubject_Change
    End If
    
End Sub

'--------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If UnloadMode = vbFormControlMenu Then
    End If
    
    Call TidyUpOnExit
    
Exit Sub
ErrHandler:
'    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "Form_QueryUnload", Err.Source) = Retry Then
'        Resume
'    End If
End Sub

'--------------------------------------------------------------------
Private Sub Form_Resize()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    'check that form has not been minimised
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If Me.ScaleWidth <= mlMinScaleWidth Then
        ' Fit to min. width if width less than minimum
        Call FitToWidth(mlMinScaleWidth)
    Else
        Call FitToWidth(Me.ScaleWidth)
    End If

    If Me.ScaleHeight <= mlMinScaleHeight Then
        ' Set to the "minimum" height
        Call FitToHeight(mlMinScaleHeight)
    Else
        Call FitToHeight(Me.ScaleHeight)
    End If

Exit Sub
ErrHandler:
'    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "Form_Resize", Err.Source) = Retry Then
'        Resume
'    End If
End Sub

'--------------------------------------------------------------------
Private Sub FitToHeight(ByVal lWinHeight As Long)
'--------------------------------------------------------------------
' Fit the controls into the given window height
' Assume the height is not below the minimum
'--------------------------------------------------------------------

    ' Move the Revalidation area down (we don't change its height)
    cmdExit.Top = lWinHeight - sbrMenu.Height - mlGAP - cmdExit.Height
    fraRevalidation.Top = cmdExit.Top - mlGAP - fraRevalidation.Height
    lblFile.Top = fraRevalidation.Top + fraRevalidation.Height + mlGAP
    
    ' Pull the subject list down to meet it
'    fraSubjects.Height = fraRevalidation.Top - mlGAP - fraSubjects.Top
    ' Size the lstSubjects first because it won't be exact
'    lstSubjects.Height = fraSubjects.Height - lstSubjects.Top - cmdRefresh.Height - 2 * mlGAP
'    cmdRefresh.Top = lstSubjects.Top + lstSubjects.Height + mlGAP
'    lblNSubjects.Top = cmdRefresh.Top
    
    ' Make sure status bar always sits on top
    sbrMenu.ZOrder
    
End Sub

'--------------------------------------------------------------------
Private Sub FitToWidth(ByVal lWinWidth As Long)
'--------------------------------------------------------------------
' Fit the controls into the given window width
' Assume the width is not below the minimum
'--------------------------------------------------------------------

    ' Expand the subject list area
'    fraSubjects.Width = lWinWidth - fraSubjects.Left - mlGAP
'    lstSubjects.Width = fraSubjects.Width - 2 * lstSubjects.Left
'    cmdRefresh.Left = lstSubjects.Left + lstSubjects.Width - cmdRefresh.Width
    
    ' Expand the revalidation area
    fraRevalidation.Width = lWinWidth - fraRevalidation.Left - mlGAP
    txtMsg.Width = fraRevalidation.Width - 2 * txtMsg.Left
    cmdDoXML.Left = txtMsg.Left + txtMsg.Width - cmdDoXML.Width
    cmdXMLInput.Left = cmdDoXML.Left - cmdXMLInput.Width - mlGAP
'    pbBar.Width = txtMsg.Width
    
    ' Finally move the Exit button over
    cmdExit.Left = fraRevalidation.Left + fraRevalidation.Width - cmdExit.Width

End Sub

'--------------------------------------------------------------------
Private Sub mnuFExit_Click()
'--------------------------------------------------------------------
    
    On Error GoTo ErrHandler

    Call TidyUpOnExit
    
    Unload Me

Exit Sub
ErrHandler:
'    If MACROErrorHandler(Me.Name, Err.Number, Err.Description, "cmdExit_Click", Err.Source) = Retry Then
'        Resume
'    End If
End Sub

'--------------------------------------------------------------------
Private Sub mnuHAboutMacro_Click()
'--------------------------------------------------------------------

End Sub

'--------------------------------------------------------------------
Private Sub mnuHUserGuide_Click()
'--------------------------------------------------------------------

End Sub

'---------------------------------------------------------------------
Private Sub TidyUpOnExit()
'---------------------------------------------------------------------
' Tidy up when exiting
'---------------------------------------------------------------------

    If msLockTokens > "" Then
        Call cmdUnlock_Click
    End If

End Sub

'--------------------------------------------
Private Sub DisplayMsg(sText As String)
'--------------------------------------------
' Display a message in the Message Window, followed by CR
'--------------------------------------------
    
    txtMsg.Text = txtMsg.Text & vbCrLf & sText

End Sub

'--------------------------------------------
Public Function LoadStudies() As Boolean
'--------------------------------------------
' Populate the Study combo with studies the user has access to
'--------------------------------------------
Dim lRow As Long
Dim vStudies As Variant
Dim oStudy As Study
Dim colStudies As Collection

    HourGlassOn
    
    cboStudy.Clear
    
    ' NCJ 18 Jun 03 - Bug 1856 - Get all studies (which doesn't check the user's Open Subject permission)
    Set colStudies = moUser.GetAllStudies
    
    ' Are there any studies?
    If colStudies.Count = 0 Then
        LoadStudies = False
    Else
        ' Add the studies to the combo
        ' and the study IDs to the ItemData array
        For Each oStudy In colStudies
            cboStudy.AddItem oStudy.StudyName
            cboStudy.ItemData(cboStudy.NewIndex) = oStudy.StudyId
        Next
        cboStudy.ListIndex = 0
        LoadStudies = True
    End If
 
    HourGlassOff

End Function

'--------------------------------------------
Public Sub LoadSites()
'--------------------------------------------
' Populate the Sites combo with sites the user has access to
' according to the chosen study
'--------------------------------------------
Dim lRow As Long
Dim vSites As Variant

Dim colSites As Collection
Dim oSite As Site

    cboSite.Clear
    msSelSite = ""
    HourGlassOn
    
    ' NCJ 18 Jun 03 - Bug 1856 - Get all sites (which doesn't check the user's Open Subject permission)
'    Set colSites = moUser.GetOpenSubjectSites(cboStudy.ItemData(cboStudy.ListIndex))
    Set colSites = moUser.GetAllSites(cboStudy.ItemData(cboStudy.ListIndex))
    Set mcolWritableSites = New Collection
    
    ' Are there any sites?
    If colSites.Count > 0 Then
        cboSite.AddItem msALL_SITES
        For Each oSite In colSites
            ' Can't open subjects from Remote sites on the Server
            If Not (moUser.DBIsServer And oSite.SiteLocation = 1) Then
                cboSite.AddItem oSite.Site
                mcolWritableSites.Add LCase(oSite.Site), LCase(oSite.Site)
            End If
        Next
    End If
    
    ' There is always the "All Sites" item
    If cboSite.ListCount > 1 Then
        cboSite.ListIndex = 0
    Else
        cboSite.Clear
        Call MsgBox("There are no subjects available for revalidation in this study")
    End If
    
    HourGlassOff
    
End Sub

'--------------------------------------------------------------------
Private Sub cboSite_Click()
'--------------------------------------------------------------------
' They clicked on a Site
'--------------------------------------------
    
    If cboSite.ListIndex > -1 Then
        If cboSite.List(cboSite.ListIndex) <> msSelSite Then
            msSelSite = cboSite.List(cboSite.ListIndex)
        End If
    Else
        If msSelSite > "" Then
            msSelSite = ""
        End If
    End If
    Call DoSubjectSpec
    
End Sub

'--------------------------------------------
Private Sub cboStudy_Click()
'--------------------------------------------
' They clicked on a Study
'--------------------------------------------

    ' Any study chosen?
    If cboStudy.ListIndex > -1 Then
        If cboStudy.List(cboStudy.ListIndex) <> msSelStudyName Then
            msSelStudyName = cboStudy.List(cboStudy.ListIndex)
            mlSelStudyId = cboStudy.ItemData(cboStudy.ListIndex)
            Call LoadSites
        End If
    Else
        If msSelStudyName > "" Then
            msSelStudyName = ""
            mlSelStudyId = 0
        End If
    End If
    Call DoSubjectSpec

End Sub

'--------------------------------------------------------------------
Private Function SelectLogFile() As String
'--------------------------------------------------------------------
' Get the user to select a file to contain the revalidation results
'--------------------------------------------------------------------
Dim sLogFile As String
Dim n As Integer

    On Error GoTo CancelOpen
    
    With CommonDialog1
        .DialogTitle = "XML File"
        .InitDir = App.Path
        .DefaultExt = "xml"
        .Filter = "XML file (*.xml)|*.xml|Text file (*.txt)|*.txt|Log file (*.log)|*.log|All files (*.*)|*.*"
        .FilterIndex = 1
        .CancelError = True
        .Flags = cdlOFNCreatePrompt + cdlOFNPathMustExist _
                    + cdlOFNOverwritePrompt + cdlOFNNoReadOnlyReturn
        .ShowSave
  
        sLogFile = .FileName
    End With

    ' Create a new empty log file
    n = FreeFile
    Open sLogFile For Output As n
    Close n
    
    SelectLogFile = sLogFile
    
CancelOpen:

End Function

'--------------------------------------------------------------------
Private Sub DoSubjectSpec()
'--------------------------------------------------------------------
' Fill in the requested subject spec
'--------------------------------------------------------------------

'    If (msSelSite > "") And (msSelSite <> msALL_SITES) And (msSubjLabel > "") Then
    If (msSubjLabel > "") Then
        lblSubject.Caption = msSelStudyName & "/" & Trim(txtSubject.Text)
        Call DoSubjectXML
        cmdDoXML.Enabled = True
        cmdXMLInput.Enabled = True
    Else
        lblSubject.Caption = ""
        txtXMLSubject.Text = ""
        cmdDoXML.Enabled = False
        cmdXMLInput.Enabled = False
    End If

End Sub

'--------------------------------------------------------------------
Private Sub txtSubject_Change()
'--------------------------------------------------------------------
' Refresh subj spec. when they change the subject label
'--------------------------------------------------------------------
Dim sSubjLabel As String

    sSubjLabel = Trim(txtSubject.Text)
    txtSubject.BackColor = vbWindowBackground
    msSubjLabel = sSubjLabel
    Call DoSubjectSpec
    
End Sub


Private Sub cmdEform_Click()

    txtXML.SelText = "  <Eform Code = 'eee' Cycle = '1'>" & vbCrLf & "  </Eform>" & vbCrLf

End Sub

Private Sub cmdQuestion_Click()

    txtXML.SelText = "   <Question Code = 'qqq' Cycle = '1'/>" & vbCrLf

End Sub

Private Sub cmdVisit_Click()

    txtXML.SelText = " <Visit Code = 'vvv' Cycle = '1'>" & vbCrLf & " </Visit>" & vbCrLf

End Sub

Private Sub DoSubjectXML()
Dim sSubjXML As String

    sSubjXML = "<MACROSubject"
    sSubjXML = sSubjXML & " Study = """ & msSelStudyName & """"
'    sSubjXML = sSubjXML & " Site = """ & msSelSite & """"
    sSubjXML = sSubjXML & " Label = """ & msSubjLabel & """"
    sSubjXML = sSubjXML & ">"
    txtXMLSubject.Text = sSubjXML

End Sub

Private Function GetTheXMLRequest() As String
Dim sXML As String
Const sXMLHEADER = "<?xml version=""1.0""?>" & vbCrLf

    If txtXMLSubject.Text > "" Then
        sXML = sXMLHEADER & txtXMLSubject.Text & vbCrLf & txtXML.Text & vbCrLf & "</MACROSubject>"
    End If
    GetTheXMLRequest = sXML
    
End Function

'---------------------------------------------------------------------
Private Sub LogToFile(ByVal sText As String, _
                    Optional nMsgType As Integer)
'---------------------------------------------------------------------
' Add text to the revalidation log file (assume initialised)
' Insert no. of tabs according to msg type (Subj, Visit, eForm etc. - see MSG_xxx constants)
'---------------------------------------------------------------------
Dim n As Integer
  
    If msLogFile = "" Then
        Debug.Print "No log file: " & sText
    Else
        n = FreeFile
        Open msLogFile For Append As n
        ' Tab in by an appropriate no. of spaces
        Print #n, Space(nMsgType * 2) & sText
        
        Close n
    End If

End Sub


'----------------------------------------------------------------------------------------'
Private Function DialogQuestion(sPrompt As String, Optional sTitle As String = "", _
    Optional bCancel As Boolean = False, Optional lOptions As Long = 0) As Integer
'----------------------------------------------------------------------------------------'
'display a question msgbox
' MLM 14/02/03: Added lOptions argument in order that the default button can be specified,
'               but could be used for other MesgBox options too.
'----------------------------------------------------------------------------------------'

    If sTitle = "" Then
        'no title- use products
        sTitle = "LISA"
    End If
    
    If bCancel Then
        DialogQuestion = MsgBox(sPrompt, vbYesNoCancel + vbQuestion + lOptions, sTitle)
    Else
        DialogQuestion = MsgBox(sPrompt, vbYesNo + vbQuestion + lOptions, sTitle)
    End If
    
End Function


'---------------------------------------------------------------------
Private Sub FormCentre(frmForm As Form, Optional frmParent As Form = Nothing)
'---------------------------------------------------------------------
'   Centre form on screen or parent form
'---------------------------------------------------------------------

    If frmForm.WindowState = vbNormal Then
        If frmParent Is Nothing Then
            frmForm.Top = (Screen.Height - frmForm.Height) \ 2
            frmForm.Left = (Screen.Width - frmForm.Width) \ 2
        Else
            With frmParent
                frmForm.Top = .Top + ((.Height - frmForm.Height) \ 2)
                frmForm.Left = .Left + ((.Width - frmForm.Width) \ 2)
            End With
        End If
    End If
    
End Sub

Private Sub HourGlassOn()

    Screen.MousePointer = vbHourglass
    
End Sub

Private Sub HourGlassOff()

    Screen.MousePointer = vbNormal
    
End Sub


