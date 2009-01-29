Attribute VB_Name = "modAutoImportMain"
'----------------------------------------------------------------------------------------'
'   File:       modAutoImportMain.bas
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   Author:     Hugo de Schepper
'   Purpose:    Control of AutoImport
'----------------------------------------------------------------------------------------'
'   Notes:
'   Contains copies of routines normally in MAIN MACRO MODULE

' Revisions:
'   TA 01/05/2001: Incorporated into MACRO
'   DPH 03/05/2002 Ordered message collection by MessageTimeStamp
'   DPH 10/06/2002 - CBBL 2.2.14.26 Fixed Subject Locks causing 'repetition' on same subject
'   DPH 13/06/2002 - AutoImport has own cab extract folder (AICabExtract) CBBL 2.2.14.26
'   NCJ 24 Dec 02 - Added processing of LF messages after successful AutoImportPRD in AutoImport
'   NCJ 3 Jan 03 - Changed Initialisations to use IMedSettings file; general updates for MACRO 3.0
'   NCJ 7 Jan 03 - Added processing of LF Rollbacks in ProcessLFMessages
'   NCJ 9 Jan 03 - LF Rollback processing moved to ImportPRD in clsExchange
'   TA 19/11/2004 issue 2448 allow multiple instances for different databases
'   TA 04112005 issue 2361: autoimport no longer fills log if subject locked
'----------------------------------------------------------------------------------------'
Option Explicit

Global Const gvsDB_CLPO = "/DB:"                     'database command line parameter option
Global Const gsSINGLE_RUN = "/AI:"

'MLM 13/09/01: Use a global instance of the new (3.0) User class
Public goUser As MACROUser

Public gnTransactionControlOn As Integer

'Public MacroADODBConnection As ADODB.Connection
Public SecurityADODBConnection As ADODB.Connection

Public gsAppPath As String

Public gsTEMP_PATH As String
Public gsIN_FOLDER_LOCATION As String
Public gsOUT_FOLDER_LOCATION As String
Public gsHTML_FORMS_LOCATION As String
Public gsDOCUMENTS_PATH As String   ' NCJ 1 Oct 99
Public gsCAB_EXTRACT_LOCATION As String 'SDM 26/01/00 SR2794
' DPH 10/04/2002 - Secure HTML Location
Public gsSECURE_HTML_LOCATION As String

Public gsMACROUserGuidePath As String

Public glSystemIdleTimeout As Long
Public glSystemIdleTimeoutCount As Long

Public Const gnFIRST_ID = 1
Public Const gnID_INCREMENT = 1

' REM 07/12/01 - used to open the new .chm Help file
Public Const HH_HELP_CONTEXT = &HF

Public Const glFormColour As Long = vbButtonFace

Public Const gsDIALOG_TITLE As String = "MACRO"

Public Const valAlpha                   As Integer = 1
Public Const valNumeric                 As Integer = 2
Public Const valSpace                   As Integer = 4
Public Const valOnlySingleQuotes        As Integer = 8
Public Const valComma                   As Integer = 16
Public Const valUnderscore              As Integer = 32
Public Const valDateSeperators          As Integer = 64
'MLM 22/04/02: Added newer constants based on MainMACROModule
Public Const valMathsOperators          As Integer = 128
Public Const valDecimalPoint            As Integer = 256

'ic 27/10/2005 clinical coding on/off
Public gbClinicalCoding As Boolean
Private Const mCCSwitch As String = "CLINICALCODING"

' REM 07/12/01: MACRO help API call
Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
    (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long

'---------------------------------------------------------------------
Private Function GetAutoImportControlParams( _
    vUploadOrder As Integer, _
    vStartStop As String, _
    vWaitInterval As Long, _
    Optional bFirstTime As Boolean = False) As Integer
'---------------------------------------------------------------------
' REVISIONS
' DPH 05/04/2002 - Added Error Handling to check DB connection is OK
' Returns 0 if OK 1 if an error has occurred
'---------------------------------------------------------------------

Dim sSQL As String
Dim rsControlList As ADODB.Recordset

On Error GoTo ErrorHandler

    sSQL = "SELECT uploadorder, startstop, waitinterval " _
           & "FROM autoimportcontrol"
    Set rsControlList = New ADODB.Recordset
    rsControlList.MaxRecords = 2
    rsControlList.Open sSQL, MacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not rsControlList.EOF Then
        vUploadOrder = rsControlList!uploadorder
        vStartStop = rsControlList!StartStop
        vWaitInterval = rsControlList!WaitInterval
        GetAutoImportControlParams = 0
    Else
        GetAutoImportControlParams = 1
    End If
    
    rsControlList.Close
    Set rsControlList = Nothing
Exit Function

ErrorHandler:
    ' if an error occurs it must be DB connection problems
    ' if not first time then put into 'PAUSE' mode
    If bFirstTime Then
        GetAutoImportControlParams = 1
    Else
        GetAutoImportControlParams = 0
        vStartStop = "PAUSE"
    End If
End Function

'---------------------------------------------------------------------
Private Sub AutoImport(bSingleRun As Boolean)
'---------------------------------------------------------------------
' REVISIONS
' DPH 11/02/2002 - Added Extra Log entries for SR4553
' DPH 15/03/2002 - Added Laboratory data to AI
' DPH 05/04/2002 - Checking for Import Errors / Create Error List
'                   Error List uses Filename_E for error Filename_L for lock
'                   Errors written to Message table locks just skipped
' DPH 03/05/2002 - Ordered SQL By MessageTimeStamp
' DPH 10/06/2002 - CBBL 2.2.14.26 Fixed Subject Locks causing 'repetition' on same subject
' DPH 14/06/2002 - CBBL 2.2.15.13 - AutoImport single run
' NCJ 9-14 Jan 2003 - Include processing of Lock/Freeze messages
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsMessageList As ADODB.Recordset
Dim oExchange As clsExchange

' New Import Error Control
Dim rsMessageErrors As ADODB.Recordset
Dim sFileName As String
Dim sSiteSubjectLab As String
Dim colErrors As New Collection
Dim AnotherMacroADODBConnection As ADODB.Connection

'AutoImportControl Parameters
Dim nUploadOrder As Integer
Dim sStartStop As String
Dim lWaitInterval As Long
Dim sLDDFile As String

Dim sPRDFileName As String

Dim lStudyId As Long

Dim i As Integer


    On Error GoTo ErrHandler
    
    InitializeMacroADODBConnection
    Set oExchange = New clsExchange
    
    ' DPH 14/06/2002 - if single run do not check DB
    If Not bSingleRun Then
        i = GetAutoImportControlParams(nUploadOrder, sStartStop, lWaitInterval, True)
        If i <> 0 Then
            sStartStop = "ERROR"
        End If
    Else
        sStartStop = "START"
        lWaitInterval = 1
    End If
    
    ' Other Connection used for collecting messages
    Set AnotherMacroADODBConnection = New ADODB.Connection
    AnotherMacroADODBConnection.Open MacroADODBConnection.ConnectionString

    
    'HDS Start loop until
    While sStartStop <> "STOP" And sStartStop <> "ERROR"
    
        Select Case sStartStop
        Case "START"
            ' NCJ 9 Jan 03 - Process any unprocessed LF messages that we can
            Call DealWithUnprocessedLFMessages(oExchange)
            
            ' DPH 10/06/2002 - Reset errors collection
            Set colErrors = New Collection
            ' DPH 05/04/2002 - Collect Import Errors
            sSQL = "SELECT * FROM Message WHERE MessageDirection = " & MessageDirection.MessageIn _
                & " AND MessageReceived = " & MessageReceived.Error
            Set rsMessageErrors = New ADODB.Recordset
            rsMessageErrors.Open sSQL, AnotherMacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
            Do While Not rsMessageErrors.EOF
                sFileName = rsMessageErrors("MessageParameters")
                sSiteSubjectLab = GetSiteStudySubjectLabFromFileName(sFileName)
                Call CollectionAddAnyway(colErrors, sSiteSubjectLab & "_E", sSiteSubjectLab)
                
                ' Get next error file
                rsMessageErrors.MoveNext
            Loop
            rsMessageErrors.Close
            Set rsMessageErrors = Nothing
            
            ' DPH 03/05/2002 - Added Order By MessageTimeStamp
            sSQL = "SELECT * FROM Message WHERE MessageDirection = " & MessageDirection.MessageIn _
                & " AND MessageReceived = " & MessageReceived.NotYetReceived _
                & " ORDER BY MessageTimeStamp"
            Set rsMessageList = New ADODB.Recordset
            'rsMessageList.MaxRecords = 2
            rsMessageList.Open sSQL, AnotherMacroADODBConnection, adOpenStatic, adLockReadOnly, adCmdText
        
            If Not rsMessageList.EOF Then
            ' there is something to be uploaded
            '   Need to start transaction at this level (not just in the AutoImportPRD function)
                'TransBegin
                ' DPH 10/06/2002 - Loop through all import records
                Do While Not rsMessageList.EOF
                    Select Case rsMessageList!MessageType
                    Case ExchangeMessageType.PatientData
                        sPRDFileName = rsMessageList!MessageParameters
                        ' DPH 05/04/2002 - Check If Previous Failure for this site/subject
                        sSiteSubjectLab = GetSiteStudySubjectLabFromFileName(sPRDFileName)
                        ' If not a file that has previously errored then Import
                        If Not CollectionMember(colErrors, sSiteSubjectLab, False) Then
                            If Not SubjectIsLocked(sSiteSubjectLab, oExchange) Then
                                'TA 04112005: only try if subject is unlocked
                                ' DPH 11/02/2002 - Check If Decode Produces a file
                                ' DPH 10/04/2002 - Now using gsSECURE_HTML_LOCATION
                                If HEXDecodeFile(gsSECURE_HTML_LOCATION & sPRDFileName, _
                                                    gsIN_FOLDER_LOCATION & "_" & sPRDFileName) Then
                                    Select Case oExchange.AutoImportPRD(gsIN_FOLDER_LOCATION & "_" & sPRDFileName)
                                    Case ExchangeError.Success
                                        ' Successful
                                        sSQL = "UPDATE Message SET MessageReceived = " & MessageReceived.Received & " WHERE MessageId = " & rsMessageList!MessageId
                                    Case ExchangeError.SubjectLock
                                        ' Add to errors collection as a lock (do nothing with message)
                                        sSQL = ""
                                        Call CollectionAddAnyway(colErrors, sSiteSubjectLab & "_L", sSiteSubjectLab)
                                    Case Else
                                        ' Set to failed
                                        sSQL = "UPDATE Message SET MessageReceived = " & MessageReceived.Error & " WHERE MessageId = " & rsMessageList!MessageId
                                        Call CollectionAddAnyway(colErrors, sSiteSubjectLab & "_E", sSiteSubjectLab)
                                    End Select
                                Else
                                    ' Write failed hex decode message to Log table
                                    gLog "AutoImport", "Hex decode of " & sPRDFileName & " failed. AutoImport Error."
                                    ' Set to failed
                                    sSQL = "UPDATE Message SET MessageReceived = " & MessageReceived.Error & " WHERE MessageId = " & rsMessageList!MessageId
                                    Call CollectionAddAnyway(colErrors, sSiteSubjectLab & "_E", sSiteSubjectLab)
                                End If
                            End If
                        Else
                            ' Set message to skipped (if previous cab in error)
                            If Right(colErrors(sSiteSubjectLab), 2) = "_E" Then
                                sSQL = "UPDATE Message SET MessageReceived = " & MessageReceived.Skipped & " WHERE MessageId = " & rsMessageList!MessageId
                            Else
                                sSQL = ""
                            End If
                            gLog "AutoImport", "Skipped file " & sPRDFileName & " due to preceding error in Subject/Lab AI file. AutoImport Warning."
                        End If
                        If sSQL <> "" Then
                            MacroADODBConnection.Execute sSQL, dbFailOnError
                        End If
                        
                    Case ExchangeMessageType.LabDefinitionSiteToServer
                         ' DPH 05/04/2002 - Check If Previous Failure for this site/subject
                        sSiteSubjectLab = GetSiteStudySubjectLabFromFileName(rsMessageList!MessageParameters)
                        ' If not a file that has previously errored then Import
                        If Not CollectionMember(colErrors, sSiteSubjectLab, False) Then
                            'UnHex the imported laboratory definition file
                            ' DPH 25/03/2002 - Check If Decode Produces a file
                            ' DPH 10/04/2002 - Now using gsSECURE_HTML_LOCATION
                            'If HEXDecodeFile(gsHTML_FORMS_LOCATION & rsMessageList!MessageParameters, gsIN_FOLDER_LOCATION & "_" & rsMessageList!MessageParameters) Then
                            If HEXDecodeFile(gsSECURE_HTML_LOCATION & rsMessageList!MessageParameters, gsIN_FOLDER_LOCATION & "_" & rsMessageList!MessageParameters) Then
                                'Unpack the CAB file into an LDD file,
                                oExchange.ImportLDDCAB gsIN_FOLDER_LOCATION & "_" & rsMessageList!MessageParameters
                                'get the name of the ldd file using the DIR command
                                sLDDFile = Dir(gsCAB_EXTRACT_LOCATION & "*.ldd")
                                If sLDDFile > "" Then
                                    Select Case oExchange.ImportLDD(gsCAB_EXTRACT_LOCATION & sLDDFile)
                                    Case ExchangeError.EmptyFile
                                        gLog "AutoImportLDD", sLDDFile & " was empty. Import aborted"
                                        sSQL = "UPDATE Message SET MessageReceived = " & MessageReceived.Error & " WHERE MessageId = " & rsMessageList!MessageId
                                        Call CollectionAddAnyway(colErrors, sSiteSubjectLab & "_E", sSiteSubjectLab)
                                    Case ExchangeError.Invalid
                                         gLog "AutoImportLDD", sLDDFile & " was not a valid laboratory definition file. Import aborted"
                                         sSQL = "UPDATE Message SET MessageReceived = " & MessageReceived.Error & " WHERE MessageId = " & rsMessageList!MessageId
                                        Call CollectionAddAnyway(colErrors, sSiteSubjectLab & "_E", sSiteSubjectLab)
                                    Case ExchangeError.Success
                                        gLog "AutoImportLDD", sLDDFile & " imported successfully"
                                        sSQL = "UPDATE Message SET MessageReceived = " & MessageReceived.Received & " WHERE MessageId = " & rsMessageList!MessageId
                                    Case Else
                                        gLog "AutoImportLDD", sLDDFile & " unexpected. Import aborted"
                                        sSQL = "UPDATE Message SET MessageReceived = " & MessageReceived.Error & " WHERE MessageId = " & rsMessageList!MessageId
                                        Call CollectionAddAnyway(colErrors, sSiteSubjectLab & "_E", sSiteSubjectLab)
                                    End Select
                                Else
                                    gLog "AutoImportLDD", "No LDD files extracted from " & StripFileNameFromPath(rsMessageList!MessageParameters) & " . AutoImport Error (Lab)."
                                    sSQL = "UPDATE Message SET MessageReceived = " & MessageReceived.Error & " WHERE MessageId = " & rsMessageList!MessageId
                                    Call CollectionAddAnyway(colErrors, sSiteSubjectLab & "_E", sSiteSubjectLab)
                                End If
                                
                                'Kill the Cab File
                                'Kill gsIN_FOLDER_LOCATION & "_" & rsMessageList!MessageParameters
                            Else
                                ' Write failed hex decode message to Log table
                                gLog "AutoImport", "Hex decode of " & StripFileNameFromPath(rsMessageList!MessageParameters) & " failed. AutoImport Error."
                                sSQL = "UPDATE Message SET MessageReceived = " & MessageReceived.Error & " WHERE MessageId = " & rsMessageList!MessageId
                            End If
                        Else
                            ' Set message to skipped
                            sSQL = "UPDATE Message SET MessageReceived = " & MessageReceived.Skipped & " WHERE MessageId = " & rsMessageList!MessageId
                            Call CollectionAddAnyway(colErrors, sSiteSubjectLab & "_E", sSiteSubjectLab)
                            gLog "AutoImport", "Skipped file " & StripFileNameFromPath(rsMessageList!MessageParameters) & " due to preceding error in Subject/Lab AI file. AutoImport Warning."
                        End If
                        
                        MacroADODBConnection.Execute sSQL, dbFailOnError
                    
                    End Select
                    
                    ' DPH 10/06/2002 - Get next record
                    rsMessageList.MoveNext
                Loop
                'TA 03112005: Wait a little bit before checking again if there is new data
                Sleep lWaitInterval * 1000
            Else
                'Wait a little bit before checking again if there is new data
                Sleep lWaitInterval * 1000
            End If
            rsMessageList.Close
            Set rsMessageList = Nothing
        Case "PAUSE"
            Sleep lWaitInterval * 1000
        End Select
        ' DPH 14/06/2002 -
        If Not bSingleRun Then
            i = GetAutoImportControlParams(nUploadOrder, sStartStop, lWaitInterval)
            If i <> 0 Then
                sStartStop = "ERROR"
            End If
        Else
            sStartStop = "STOP"
        End If
    Wend
    
    AnotherMacroADODBConnection.Close
    Set AnotherMacroADODBConnection = Nothing

    Exit Sub
    
ErrHandler:
        'TA 22/08/2005    write error to file (may be db error)
    WriteLogError "modAutoImportMain.AutoImport", "Error occurred during import. Error code " & Err.Number & " - " & Err.Description

End Sub

'---------------------------------------------------------------------
Sub Main()
'---------------------------------------------------------------------
' REVISIONS
' DPH 11/06/2002 - CBBL 2.2.14.15 - Stop 2 instances of MACRO_AI simultaneously
'   running on same machine
' DPH 14/06/2002 - CBBL 2.2.15.13 - AutoImport single run
' NCJ 3 Jan 03 - Got this routine working for MACRO 3.0
' TA 19/11/2004 issue 2448 allow multiple instances for different databases
'---------------------------------------------------------------------
Dim sDatabase As String
Dim sSecCon As String
Dim bAllowMultiAI As Boolean
Dim nLockFile As Integer
Dim sLockFile As String
#If rochepatch <> 1 Then
Dim oVersion As MACROVersion.Checker
#End If
    ' exit if not correctly entered parameter
    If (UCase(Left(Command, 4)) <> gvsDB_CLPO) And (UCase(Left(Command, 4)) <> gsSINGLE_RUN) Then
        End
    End If
       
    sDatabase = Mid(Command, Len(gvsDB_CLPO) + 1)
    If sDatabase = "" Then
        End
    End If
     
    ' NCJ 3 Jan 2003 - Initialise settings file
    Call InitialiseSettingsFile
    'Get initialisationsettings from registry
    Call InitialisationSettings
    
    'are they allowed to run more than one
    bAllowMultiAI = (LCase(GetMACROSetting("allowmultiai", "false")) = "true")

    'TA 19/11/2004 issue 2448 allow multiple instances for different databases
    If bAllowMultiAI Then
        'create a lock file so only one AI can run per db
        sLockFile = gsCAB_EXTRACT_LOCATION & sDatabase & ".lock"
        nLockFile = FreeFile
        On Error Resume Next
        Open sLockFile For Output Lock Read Write As #nLockFile
        If Err.Number <> 0 Then
            'it must be locked
            MsgBox "Auto Import already running for " & sDatabase
            Exit Sub
        End If
        On Error GoTo 0
        'if it is multi use a subfolder of the normal cab extract named after database
        gsCAB_EXTRACT_LOCATION = gsCAB_EXTRACT_LOCATION & sDatabase & "\"
    Else
        If App.PrevInstance Then
            ' DPH 14/06/2002 - CBBL 2.2.15.13 - AutoImport single run
            MsgBox "Auto Import already running"
            Exit Sub
        End If
    End If
        
    ' To reach here AI is not already running or multiple running is allowed
        
    gnTransactionControlOn = 0
    'Changed by Mo Morris 28/10/99 - ADO InitializeSecurityADODBConnection call added
    sSecCon = InitializeSecurityADODBConnection
    Set goUser = SilentUserLogin(sSecCon, sDatabase, gsAppPath & "HTML")
    If goUser Is Nothing Then
        ExitMACRO
        MACROEnd
    End If

    ' NCJ 3 Jan 03 - Get Secure HTML location from User object
    gsSECURE_HTML_LOCATION = goUser.Database.SecureHTMLLocation
    InitializeMacroADODBConnection
    
#If rochepatch <> 1 Then
    'check for clinical coding version
    Set oVersion = New MACROVersion.Checker
    gbClinicalCoding = oVersion.HasUpgrade(goUser.CurrentDBConString, mCCSwitch)
    Set oVersion = Nothing
#End If
    ' RunAutoImport version required, automatic or single run
    If UCase(Left(Command, 4)) = gvsDB_CLPO Then
        ' automatic - database controlled
        Call AutoImport(False)
    Else
        ' single run
        Call AutoImport(True)
    End If


    If bAllowMultiAI Then
        'clean up lock file
        Close #nLockFile
        On Error Resume Next
        Kill sLockFile
        On Error GoTo 0
    End If

End Sub

'---------------------------------------------------------------------
Private Sub InitialisationSettings()
'---------------------------------------------------------------------
' REVISIONS
' DPH 13/06/2002 - AutoImport has own cab extract folder (AICabExtract) CBBL 2.2.14.26
' NCJ 3 Jan 03 - Use new IMEDSettings component to get settings
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    'set-up a Application Path variable
    gsAppPath = App.Path
                                
    AddDirSep gsAppPath
    
    'set-up specific Path variables for application folders
    ' NB every PATH or LOCATION variable ends with a backslash
     gsTEMP_PATH = GetMACROSetting("Temp", gsAppPath & "Temp\")
    
    ' DPH 18/10/2001 - Need to check for TEMP folders existence - vital to AREZZO
    ' Only Applicable on MACRO_DM, MACRO_SD - Where Arezzo is used
    ' Quit if no Temp Folder
    If App.Title = "MACRO_SD" Or App.Title = "MACRO_DM" Then
        If Not (FolderExistence(gsTEMP_PATH & "dummy.txt")) Then
            ' Show User error Message
            Call DialogError("Temporary Path " & gsTEMP_PATH & vbCrLf & "could not be created or has no write permissions." & vbCrLf & "MACRO cannot work without this folder")
            ' Quit
            End
        End If
    End If
    
    ' NCJ 3 Jan 03 - Use new IMEDSettings component to get named settings
    ' or return the default passed
    gsIN_FOLDER_LOCATION = GetMACROSetting("In Folder", gsAppPath & "In Folder\")
    gsOUT_FOLDER_LOCATION = GetMACROSetting("Out Folder", gsAppPath & "Out Folder\")
    gsDOCUMENTS_PATH = GetMACROSetting("Documents", gsAppPath & "Documents\")
    gsCAB_EXTRACT_LOCATION = GetMACROSetting("CabExtract", gsAppPath & "AICabExtract\")
    gsMACROUserGuidePath = GetMACROSetting("Help", gsAppPath & "Help\")
        
    Exit Sub
    
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                    "InitialisationSettings", "modAutoImportMain")
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
Private Function SilentUserLogin(sSecCon As String, ByVal sDatabase As String, _
                                    sDefaultHTMLLocation As String) As MACROUser
'---------------------------------------------------------------------
' NCJ 3 Jan 03 - Replaced with new MACRO 3.0 version
'---------------------------------------------------------------------
Dim oUser As MACROUser
Dim bLoginSucceeded As Boolean
Dim sMessage As String

    Set oUser = New MACROUser
    bLoginSucceeded = oUser.SilentLogin(sSecCon, sDatabase, sDefaultHTMLLocation, sMessage)
    
    If bLoginSucceeded Then
        Set SilentUserLogin = oUser
    Else
        Set SilentUserLogin = Nothing
    End If

End Function


'---------------------------------------------------------------------
Public Function MACROFormErrorHandler(oForm As Form, nTrappedErrNum As Long, _
        sTrappedErrDesc As String, sProcName As String) As OnErrorAction
'---------------------------------------------------------------------
' Call the error handling form pass it the err no. the err desc and the
' form that raised the error.
'---------------------------------------------------------------------
    On Error GoTo MemoryErr
    
    'za 28/09/01
    'here we call the new error handler routine
    MACROFormErrorHandler = MACROErrorHandler(oForm.Name, nTrappedErrNum, sTrappedErrDesc, sProcName, GetApplicationTitle)
    
   
' SR3685 Trap for out of memory
    
Exit Function
MemoryErr:
    Select Case Err.Number
        Case 7 ' Out of Memory error
            MsgBox "The application has run out of memory and will now be shut down.", vbCritical, "MACRO"
            Call ExitMACRO
            Call MACROEnd
    End Select

End Function

'---------------------------------------------------------------------
Public Function MACROCodeErrorHandler(nTrappedErrNum As Long, sTrappedErrDesc As String, _
                            sProcName As String, sModuleName As String) As OnErrorAction
'---------------------------------------------------------------------
' Call the error handling form pass it the err no. the err desc
' there is no form to pass as this one is used in the modules.
'---------------------------------------------------------------------
    On Error GoTo MemoryErr
    
    'za 28/09/01
    'here we call the new error handler routine
    MACROCodeErrorHandler = MACROErrorHandler(sModuleName, nTrappedErrNum, sTrappedErrDesc, sProcName, GetApplicationTitle)
    
' SR3685 Trap for out of memory
Exit Function
MemoryErr:
    Select Case Err.Number
        Case 7 ' Out of Memory error
            MsgBox "The application has run out of memory and will now be shut down.", vbCritical, "MACRO"
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Function

'---------------------------------------------------------------------
Public Sub ExitMACRO()
'---------------------------------------------------------------------
Dim mDatabase As Database
    
    'Close all open files
    Close
    
    ' PN 23/09/99
    ' ensure that the ado connection is terminated properly
    Call TerminateAllADODBConnections
    
End Sub

'---------------------------------------------------------------------
Public Sub MACROEnd()
'---------------------------------------------------------------------
' NCJ 13 Dec 99
' End statement for MACRO
' (This should not be done if compiling a DLL)
'---------------------------------------------------------------------

    End
    
End Sub

'----------------------------------------------------------------------------------------'
Public Function GetApplicationTitle() As String
'----------------------------------------------------------------------------------------'
'Return the default title of an app
'----------------------------------------------------------------------------------------'

    Select Case App.Title
    Case "MACRO_SD"
        If LCase$(Command) = "library" Then
            GetApplicationTitle = "MACRO Library Management"
        Else
            GetApplicationTitle = "MACRO Study Definition"
        End If
    Case "MACRO_DM"
         If LCase$(Command) = "review" Then
            GetApplicationTitle = "MACRO Data Review"
        Else
            'TA 23/10/2000 changend from Data Management"
            GetApplicationTitle = "MACRO Data Entry"
        End If
    Case "MACRO_SM"
        GetApplicationTitle = "MACRO System Management"
    Case "MacroCreateDataViews"
        GetApplicationTitle = "MACRO"
    Case Else
        'TA 17/1/02: part of VTRACK buglist build 1.0.3 Bug 6
        'else case just returns the real app.title
        GetApplicationTitle = App.Title
    End Select
    
End Function

'---------------------------------------------------------------------
Public Sub MACROHelp(ByVal lhWnd As Long, ByVal AppTitle As String)
'---------------------------------------------------------------------
' REM 07/12/01
' Opens MACRO help based on the context ID of the different MACRO modules
'REVISIONS:
' REM 17/01/02 - added Case MacroCreateDataViews as it did not exist
'---------------------------------------------------------------------
Dim hwndHelp

    'Context ID's for each module in MACRO Help
Const lDATAENTRY As Long = 1
Const lDATAREVIEW As Long = 2
Const lWEBDEDR As Long = 3
Const lLIBRARYMANAGEMENT = 4
Const lSTUDYDEFINITION = 5
Const lEXCHANGE = 6
Const lSYSTEMMANAGEMENT = 7
Const lCREATEDATAVIEWS = 8
Const lMACROWELCOME = 9
    
    ' Calls the help which is now contained in the MACRo.chm file that
    ' requires a context ID to open the specific help for each module
    
    Select Case AppTitle
    Case "MACRO_SD"
        If LCase$(Command) = "library" Then
            hwndHelp = HtmlHelp(lhWnd, gsAppPath & "help\MACRO.chm", HH_HELP_CONTEXT, lLIBRARYMANAGEMENT)
        Else
            hwndHelp = HtmlHelp(lhWnd, gsAppPath & "help\MACRO.chm", HH_HELP_CONTEXT, lSTUDYDEFINITION)
        End If
    Case "MACRO_DM"
         If LCase$(Command) = "review" Then
            hwndHelp = HtmlHelp(lhWnd, gsAppPath & "help\MACRO.chm", HH_HELP_CONTEXT, lDATAREVIEW)
        Else
            hwndHelp = HtmlHelp(lhWnd, gsAppPath & "help\MACRO.chm", HH_HELP_CONTEXT, lDATAENTRY)
        End If
    Case "MACRO_SM"
        hwndHelp = HtmlHelp(lhWnd, gsAppPath & "help\MACRO.chm", HH_HELP_CONTEXT, lSYSTEMMANAGEMENT)
    'REM 17/01/02 - Added MacroCteateDataViews to Help
    Case "MacroCreateDataViews"
        hwndHelp = HtmlHelp(lhWnd, gsAppPath & "help\MACRO.chm", HH_HELP_CONTEXT, lCREATEDATAVIEWS)
    Case Else
        hwndHelp = HtmlHelp(lhWnd, gsAppPath & "help\MACRO.chm", HH_HELP_CONTEXT, lMACROWELCOME)
    End Select

    
End Sub

'----------------------------------------------------------------------------------------'
Public Property Get SecurityDatabasePath() As String
'----------------------------------------------------------------------------------------'

    SecurityDatabasePath = GetMACROSetting("SecurityPath", DefaultSecurityDatabasePath)
    If SecurityDatabasePath <> "" Then
        SecurityDatabasePath = DecryptString(SecurityDatabasePath)
    End If
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Let SecurityDatabasePath(sNewSecurityPath As String)
'----------------------------------------------------------------------------------------'
' NCJ 24/1/00 - Allow setting of default path (e.g. from System Management)
' Assume sNewSecurityPath is valid
' Note this does NOT change the local value
' ASH 12/9/2002 - Registry keys replaced with calls to new Settings file
'----------------------------------------------------------------------------------------'
Dim sRegPath As String
    
    'NCJ 3 Jan 2003 - Use MACRO 3.0 new IMEDSettings component to add to settings file
    Call SetMACROSetting("SecurityPath", EncryptString(sNewSecurityPath))
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get DefaultSecurityDatabasePath() As String
'----------------------------------------------------------------------------------------'
' Get MACRO's default security path (i.e. the one set on installation)
'----------------------------------------------------------------------------------------'

    DefaultSecurityDatabasePath = App.Path & ""

End Property

'----------------------------------------------------------------------------------------'
Private Sub DealWithUnprocessedLFMessages(oExchange As clsExchange)
'----------------------------------------------------------------------------------------'
' NCJ 14 Jan 03 - Deal with all the unprocessed LF messages
' for subjects which have no data import files
' This includes doing all the rollbacks that we can too
'----------------------------------------------------------------------------------------'
Dim colIgnoreSubjects As Collection
Dim sSQL As String
Dim rsImports As ADODB.Recordset
Dim oLockFreezer As LockFreeze
Dim sKey As String
Dim sPrevKey As String
Dim oLFMsg As LFMessage
Dim oPrevLFMsg As LFMessage
Dim sLockToken As String

    On Error GoTo ErrHandler
    ' We look for LF messages for any subjects that don't have any data files waiting for import
    ' Get a collection of subjects which have data files awaiting
    sSQL = "SELECT ClinicalTrialName, TrialSite, PersonId FROM DataImport " _
            & " ORDER BY ClinicalTrialName, TrialSite, PersonId"
    Set rsImports = New ADODB.Recordset
    rsImports.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockPessimistic, adCmdText
    
    Set colIgnoreSubjects = New Collection
    ' Create collection of subjects from the recordset
    If rsImports.RecordCount > 0 Then
        rsImports.MoveFirst
        Do While Not rsImports.EOF
            sKey = rsImports!ClinicalTrialName & "|" & rsImports!TrialSite & "|" & rsImports!PersonID
            ' Add a dummy item of "1" with specified key
            Call CollectionAddAnyway(colIgnoreSubjects, 1, sKey)
            rsImports.MoveNext
        Loop
    End If

    rsImports.Close
    Set rsImports = Nothing
    
    Set oLockFreezer = New LockFreeze
    sPrevKey = ""
    sLockToken = ""
    Set oPrevLFMsg = Nothing
    
    ' Now process the LF Messages for subjects NOT in this collection
    For Each oLFMsg In oLockFreezer.ProcessableMessages(MacroADODBConnection, colIgnoreSubjects)
        sKey = oLFMsg.StudyId & "|" & oLFMsg.Site & "|" & oLFMsg.SubjectId
        If sKey <> sPrevKey Then
            ' NEW SUBJECT
            ' Need to unlock prev. subject and lock this one
            If sLockToken > "" And Not (oPrevLFMsg Is Nothing) Then
                Call oExchange.RemoveSubjectLock(oPrevLFMsg.StudyId, oPrevLFMsg.Site, oPrevLFMsg.SubjectId, sLockToken)
            End If
            
            ' Remember this subject
            sPrevKey = sKey
            Set oPrevLFMsg = oLFMsg
            
            ' Try for a lock on the new subject
            sLockToken = oExchange.GetSubjectLock(goUser.UserName, oLFMsg.StudyId, oLFMsg.Site, oLFMsg.SubjectId)
            
            ' Must hoover up any Rollbacks before doing the new ones
            If sLockToken > "" Then
                gLog "AutoImport", "Handling Lock/Freeze messages for site " & oLFMsg.Site & ", subject " & oLFMsg.SubjectId
                Call oLockFreezer.ProcessSubjectRollBacks(MacroADODBConnection, oLFMsg.StudyName, oLFMsg.Site, oLFMsg.SubjectId)
            End If
        End If
        
        ' Only do the operation if we have a lock on the subject
        ' (Otherwise we ignore this message and leave it until the next AutoImport session)
        If sLockToken > "" Then
            Call oLFMsg.DoAction(MacroADODBConnection, True)
        End If
    Next
    
    ' Finally remove the subject lock for the last message (if there was one)
    If sLockToken > "" Then
        Call oExchange.RemoveSubjectLock(oPrevLFMsg.StudyId, oPrevLFMsg.Site, oPrevLFMsg.SubjectId, sLockToken)
    End If
    
    ' Now go and do the Rollbacks similarly for subjects not already dealt with
    Call ProcessRollbackMessages(oExchange)
    
    Set oLockFreezer = Nothing
    Set colIgnoreSubjects = Nothing
    Set oLFMsg = Nothing
    Set oPrevLFMsg = Nothing
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modAutoImportMain.DealWithUnprocessedLFMessages"
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub ProcessRollbackMessages(oExchange As clsExchange)
'----------------------------------------------------------------------------------------'
' NCJ 14 Jan 03 - Deal with all the remaining unprocessed LF messages that we can
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim oLockFreezer As LockFreeze
Dim sKey As String
Dim sPrevKey As String
Dim oLFMsg As LFMessage
Dim oPrevLFMsg As LFMessage
Dim sLockToken As String
    
    On Error GoTo ErrHandler
    
    Set oLockFreezer = New LockFreeze
    sPrevKey = ""
    sLockToken = ""
    Set oPrevLFMsg = Nothing
    
    ' Process the rollback messages not dealt with so far
    For Each oLFMsg In oLockFreezer.RollbackMessages(MacroADODBConnection)
        sKey = oLFMsg.StudyId & "|" & oLFMsg.Site & "|" & oLFMsg.SubjectId
        If sKey <> sPrevKey Then
            ' NEW SUBJECT
            ' Need to unlock prev. subject and lock this one
            If sLockToken > "" And Not (oPrevLFMsg Is Nothing) Then
                Call oExchange.RemoveSubjectLock(oPrevLFMsg.StudyId, oPrevLFMsg.Site, oPrevLFMsg.SubjectId, sLockToken)
            End If
            
            ' Remember this subject
            sPrevKey = sKey
            Set oPrevLFMsg = oLFMsg
            
            ' Try for a lock on the new subject
            sLockToken = oExchange.GetSubjectLock(goUser.UserName, oLFMsg.StudyId, oLFMsg.Site, oLFMsg.SubjectId)
            If sLockToken > "" Then
                gLog "AutoImport", "Handling Lock/Freeze rollback messages for site " & oLFMsg.Site & ", subject " & oLFMsg.SubjectId
            End If
        End If
        
        ' Only do the operation if we have a lock on the subject
        ' (Otherwise we ignore this message and leave it until the next AutoImport session)
        If sLockToken > "" Then
            Call oLockFreezer.HandleRollBack(MacroADODBConnection, oLFMsg)
        End If
    Next
    
    ' Finally remove the subject lock for the last message (if there was one)
    If sLockToken > "" Then
        Call oExchange.RemoveSubjectLock(oPrevLFMsg.StudyId, oPrevLFMsg.Site, oPrevLFMsg.SubjectId, sLockToken)
    End If
    
    Set oLockFreezer = Nothing
    Set oLFMsg = Nothing
    Set oPrevLFMsg = Nothing
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modAutoImportMain.ProcessRollbackMessages"
    
End Sub

'----------------------------------------------------------------------------------------'
Private Function SubjectIsLocked(sSitePerson As String, oExchange As clsExchange) As Boolean
'----------------------------------------------------------------------------------------'
'TA 04112005 - routine that reports whether subject is locked
'----------------------------------------------------------------------------------------'
'sSitePerson in the form: sSite & "_" & sTrialName & "_" & lPersonId
Dim sToken As String
Dim sSite As String
Dim lPersonId As Long
Dim lStudyId As Long
Dim vSiteStudyPerson As Variant
Dim sStudy As String
Dim i As Long

    vSiteStudyPerson = Split(sSitePerson, "_")
    sSite = vSiteStudyPerson(0)
    lPersonId = CLng(vSiteStudyPerson(UBound(vSiteStudyPerson)))
    For i = 1 To UBound(vSiteStudyPerson) - 2
        sStudy = sStudy & vSiteStudyPerson(i) & "_"
    Next
    sStudy = sStudy & vSiteStudyPerson(UBound(vSiteStudyPerson) - 1)
    
    lStudyId = TrialIdFromName(sStudy)
   
    sToken = oExchange.GetSubjectLock(goUser.UserName, lStudyId, sSite, lPersonId)
    If sToken = "" Then
        SubjectIsLocked = True
    Else
        Call oExchange.RemoveSubjectLock(lStudyId, sSite, lPersonId, sToken)
        SubjectIsLocked = False
    End If
    Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|modAutoImportMain.SubjectIsLocked"
        
End Function
