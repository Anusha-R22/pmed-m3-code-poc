Attribute VB_Name = "modDataEntry"
'----------------------------------------------------------------------------------------'
'   File:       modDataEntry.bas
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Nicky Johns, May 2000
'   Purpose:    Routines for support of MACRO Data Entry
'----------------------------------------------------------------------------------------'
' Revisions:
'   NCJ 30 May 2000 - Created (only ShowEForm)
'   NCJ 20 June 2000 - Code and declarations for Study Object collections
'   TA 30/06/2000 - Routines for handling the study object collections
'                    moved to modStudyObjects from modDataEntry 30/06/2000)
' NCJ 19/10/00 - Assign Lab Code to eForm after building form
' NCJ 31/1/01 - Call SetFormContext in ShowEForm
' TA LoadSubject rewritten
' TA 2/10/01: Locking fixed (hopefully)
' TA 2/10/01: Better error handling
' TA 3/10/01: Fixed problem of being able to open a subject that was already locked
' TA 9/10/01: Fixed problem of new subject error when the study def is being edited
' TA 13/11/01: Functions changed so that Arezzo, StudyDef and SubjectToken are passed in and out to make this module less dependant
' TA 22/11/01: Changes to avoid errors when a studydef object is loaded without a subject object
' NCJ 22 Mar 02 - Added UpdateMode to LoadSubject calls
' TA 13/05/02: Added multiuser support - note this only written for Oracle
'       if MUMode = 1 then new version of LockSubject and UnlockSubject are compiled
' TA 22/05/02: New subject dummy lock uses old style locking now in MU mode
'TA CBB 2.2.20.3 18/7/02: only insert a row into the arezzo table if we preload a study
'TA 22/10/2002: Added more error handlers
'TA 06/12/2002: Reinstated study preloading
' NCJ 14 Jan 03 - Fixed error message in LockSubject
' RS 20 Jan 03 - Set Local Date & Time format in LoadSubject
' NCJ 23 Jan 03 - Only set LocalDateFormat in LoadSubject if user wants to use it
' NCJ 24 Jan 03 - Added sCountry parameter to NewSubject
' NCJ 6 May 03 - Added UserNameFull and UserRole to New/Load subject
' NCJ 22 Aug 06 - Use LOCK DLL for cache handling; sorted out cache issues!
'----------------------------------------------------------------------------------------'

Option Explicit

'TA 20/9/01: Our 2 public objects required for Data Management
Public goArezzo As Arezzo_DM
Public goStudyDef As StudyDefRO

'REM 24/04/02 -
Private msArezzoId As String
' NCJ 22 Aug 06 - This is the study to which the token relates
Private mlStudyId As Long

Public Enum eWinForms
    wfHome
    wfNewSubject
    wfOpenSubject
    wfSubject
    wfDiscepancies
    wfNotes
    wfSDV
    wfDataBrowser
    wfBlank
End Enum


Public gcolTopFormOrder As Collection
Public gcolBottomFormOrder As Collection

'---------------------------------------------------------------------
Public Sub EnsureCaptionsCorrect()
'---------------------------------------------------------------------
'resets the caption of the top border to what it should be
'---------------------------------------------------------------------
Dim enWinForms As eWinForms
Dim i As Long

    'ensure bottom caption correct by setting the caption to the last opened form if bottom form type
    Select Case gcolTopFormOrder(gcolTopFormOrder.Count)
    Case wfDiscepancies, wfNotes, wfSDV, wfDataBrowser
        OpenWinForm gcolTopFormOrder(gcolTopFormOrder.Count)
    End Select

    'ensure top catpion correct my finding last opened form that is a top form
    For i = gcolTopFormOrder.Count To 1 Step -1
        Select Case gcolTopFormOrder(i)
        Case wfHome, wfNewSubject, wfOpenSubject, wfSubject
            enWinForms = gcolTopFormOrder(i)
            'this exit will happen as one of the above forms must be open
            Exit For
        End Select
    Next
    OpenWinForm enWinForms

End Sub

'---------------------------------------------------------------------
Public Sub OpenWinForm(enWinForms As eWinForms)
'---------------------------------------------------------------------
'takes care of setting the captions when an form is opened
'---------------------------------------------------------------------

    On Error GoTo ErrLabel

    If gcolTopFormOrder Is Nothing Then
        Set gcolTopFormOrder = New Collection
        gcolTopFormOrder.Add wfBlank, "k" & wfBlank
    End If
    
    If gcolBottomFormOrder Is Nothing Then
        Set gcolBottomFormOrder = New Collection
        gcolBottomFormOrder.Add wfBlank, "K" & wfBlank
    End If
        
    If frmMenu.SplitScreen Then
        'split screen
        
        Select Case enWinForms
        Case wfDiscepancies, wfNotes, wfSDV
            On Error Resume Next

            gcolBottomFormOrder.Remove "k" & wfDiscepancies
            gcolBottomFormOrder.Remove "k" & wfNotes
            gcolBottomFormOrder.Remove "k" & wfSDV
            
            On Error GoTo ErrLabel
            
            gcolBottomFormOrder.Add enWinForms, "k" & enWinForms
            frmMenu.moFrmBorderBottom.SetBorderCaption enWinForms
            
        Case wfDataBrowser
             On Error Resume Next
             gcolBottomFormOrder.Remove "k" & enWinForms
             On Error GoTo ErrLabel
        
            gcolBottomFormOrder.Add enWinForms, "k" & enWinForms
            frmMenu.moFrmBorderBottom.SetBorderCaption enWinForms
        Case Else
            On Error Resume Next
            
             gcolTopFormOrder.Remove "k" & enWinForms
            
             On Error GoTo ErrLabel
            gcolTopFormOrder.Add enWinForms, "k" & enWinForms
            frmMenu.mofrmBorderTop.SetBorderCaption enWinForms
        End Select

    Else
            
        On Error Resume Next
        
        'no split screen
        Select Case enWinForms
        'remove previous entries
        Case wfDiscepancies, wfNotes, wfSDV
                gcolTopFormOrder.Remove "k" & wfDiscepancies
                gcolTopFormOrder.Remove "k" & wfNotes
                gcolTopFormOrder.Remove "k" & wfSDV
        Case Else
                gcolTopFormOrder.Remove "k" & enWinForms
        End Select
        
        On Error GoTo ErrLabel
        
        gcolTopFormOrder.Add enWinForms, "k" & enWinForms
        frmMenu.mofrmBorderTop.SetBorderCaption enWinForms
    End If

    Exit Sub
    
ErrLabel:
    Err.Raise Err.no, , Err.descritpion & "|modDataEntry.OpenWinForm"
End Sub

'---------------------------------------------------------------------
Public Function CloseWinForm(enWinForms As eWinForms) As eWinForms
'---------------------------------------------------------------------
'takes care of setting the captions when an form is closed
'---------------------------------------------------------------------
Dim i As Long
    If frmMenu.SplitScreen Then
        'split screen
        Select Case enWinForms
        Case wfDiscepancies, wfNotes, wfSDV, wfDataBrowser
            On Error Resume Next
                gcolBottomFormOrder.Remove "k" & enWinForms
            On Error GoTo ErrLabel
            frmMenu.moFrmBorderBottom.SetBorderCaption gcolBottomFormOrder.Item(gcolBottomFormOrder.Count)
        Case Else
            On Error Resume Next
                gcolTopFormOrder.Remove "k" & enWinForms
            On Error GoTo ErrLabel
            frmMenu.mofrmBorderTop.SetBorderCaption gcolTopFormOrder.Item(gcolTopFormOrder.Count)
        End Select
        
    Else
        'no split screen
        On Error Resume Next
            gcolTopFormOrder.Remove "k" & enWinForms
        On Error GoTo ErrLabel
        
        frmMenu.mofrmBorderTop.SetBorderCaption gcolTopFormOrder.Item(gcolTopFormOrder.Count)
    
    End If
    
    Exit Function

ErrLabel:
    Err.Raise Err.no, , Err.descritpion & "|modDataEntry.CloseWinForm"

End Function

'---------------------------------------------------------------------
Public Function GetSubjectString(oSubject As StudySubject) As String
'---------------------------------------------------------------------
'return study/site/subject string if loaded
'---------------------------------------------------------------------
    
    GetSubjectString = ""
    If oSubject Is Nothing Then
        GetSubjectString = ""
    Else
        If oSubject.label = "" Then
            GetSubjectString = oSubject.StudyCode & "/" & oSubject.Site & "/(" & oSubject.PersonId & ")"
        Else
            GetSubjectString = oSubject.StudyCode & "/" & oSubject.Site & "/" & oSubject.label
        End If
    End If

    
End Function

'---------------------------------------------------------------------
Public Function CloseSubject(ByRef oStudyDef As StudyDefRO, _
                            bPrompt As Boolean, bUnloadStudyToo As Boolean, bCheckEformOpen As Boolean) As Boolean
'---------------------------------------------------------------------
' all subject closing to be done through here
'remove study and subject from memory
'calls frmMenu.subject close to unload forms
'Returns true if successful
'NB If user chooses not to unload study - the study definition
'   could become out of date as it is no longer locked through
'   the subject lock
'If the user unloads the study the studydef object gets set to nothing
'---------------------------------------------------------------------

    On Error GoTo Errorlabel
    If frmMenu.GUISubjectClose(bPrompt, bCheckEformOpen) Then
        If Not oStudyDef Is Nothing Then
            'TA 21/11/01: Put in check to see if  subject is loaded
            If Not oStudyDef.Subject Is Nothing Then
            
                'close registration
                Call CloseRegistration
                Call oStudyDef.RemoveSubject
            End If
            
            'always store when subject closed
            'put thing in reg
            Call StoreStudyIDinReg(oStudyDef.StudyId)
            
            If bUnloadStudyToo Then

#If VTRACK <> 1 Then
                ' NCJ 22 Aug 06 - Don't need StudyID
                Call DeleteArezzoToken
#End If
                Call oStudyDef.Terminate
                Set oStudyDef = Nothing
            End If
            
        End If
        'enable/disable taskteims
        frmMenu.EnableDisableTaskListItems

        CloseSubject = True

    Else
   
        CloseSubject = False
    End If
    Exit Function
    
Errorlabel:
    If MACROErrorHandler("modDataEntry", Err.Number, Err.Description, "CloseSubject", Err.Source) = Retry Then
        Resume
    End If
    
End Function

'---------------------------------------------------------------------
Public Function LockSubject(sUser As String, lStudyId As Long, sSite As String, lSubjectId As Long) As String
'---------------------------------------------------------------------
' Lock a subject.
' Returns if token if lock successful or empty string if not
'---------------------------------------------------------------------
Dim sLockDetails As String
Dim sMsg As String
Dim sToken As String

    On Error GoTo Errorlabel
    
    'TA 04.07.2001: use new locking
    sToken = MACROLOCKBS30.LockSubject(gsADOConnectString, sUser, lStudyId, sSite, lSubjectId)
    Select Case sToken
    Case MACROLOCKBS30.DBLocked.dblStudy
        sLockDetails = MACROLOCKBS30.LockDetailsStudy(gsADOConnectString, lStudyId)
        If sLockDetails = "" Then
            sMsg = "This study is currently being edited by another user."
        Else
            sMsg = "This study is currently being edited by " & Split(sLockDetails, "|")(0) & "."
        End If
        Call DialogInformation(sMsg)
        sToken = ""
    Case MACROLOCKBS30.DBLocked.dblSubject

        sLockDetails = MACROLOCKBS30.LockDetailsSubject(gsADOConnectString, lStudyId, sSite, lSubjectId)
        If sLockDetails = "" Then
            sMsg = "This subject is currently being edited by another user."
        Else
            sMsg = "This subject is currently being edited by " & Split(sLockDetails, "|")(0) & "."
        End If
        Call DialogInformation(sMsg)
        sToken = ""
    Case MACROLOCKBS30.DBLocked.dblEFormInstance
        ' NCJ 14 Jan 03 - Bug fix to lock message
        ' An eForm is in use, but we don't know which one, so give a generic message
        sMsg = "This subject is currently being edited by another user."
        
'        sMsg = "This subject is currently being edited."
'        sMsg = "This subject is currently being edited by " & Split(sLockDetails, "|")(0) & "."
    
        Call DialogInformation(sMsg)
        sToken = ""
    Case Else
        'hurrah, we have a lock

    End Select
    LockSubject = sToken
    Exit Function
    
Errorlabel:
        Err.Raise Err.Number, , Err.Description & "|" & "modDataEntry.LockSubject"
    Exit Function

End Function

'---------------------------------------------------------------------
Public Sub UnlockSubject(lStudyId As Long, sSite As String, lSubjectId As Long, sToken As String)
'---------------------------------------------------------------------
' Unlock the subject
'---------------------------------------------------------------------

    On Error GoTo Errorlabel
    'TA 04.07.2001: use new locking model
    If sToken <> "" Then
        'if no gsStudyToken then UnlockSubject is being called without a corresponding LockSubject being called first
        MACROLOCKBS30.UnlockSubject gsADOConnectString, sToken, lStudyId, sSite, lSubjectId
        'always set this to empty string for same reason as above
        sToken = ""
    End If
    Exit Sub
    
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modDataEntry.UnlockSubject"
   
End Sub

'----------------------------------------------------------------------
Private Sub InsertArezzoToken(lStudyId As Long)
'----------------------------------------------------------------------
' REM 24/04/02
' This will only work with one studydef object in the application
' NCJ 22 Aug 06 - Use LOCK DLL to do this, and store associated study ID
'----------------------------------------------------------------------
'Dim rsMaxDBToken As ADODB.Recordset
'Dim sArezzoId As String
'Dim lDBToken As Long
'Dim sSQL As String

    On Error GoTo Errorlabel

    msArezzoId = MACROLOCKBS30.CacheAddStudyRow(goUser.CurrentDBConString, lStudyId)
    mlStudyId = lStudyId
    
'    sArezzoId = Token
'
'    sSQL = "SELECT Max(DBToken) as MaxDBToken FROM ArezzoToken"
'    Set rsMaxDBToken = New ADODB.Recordset
'    rsMaxDBToken.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    If IsNull(rsMaxDBToken!MaxDBToken) Then
'        lDBToken = 1
'    Else
'        lDBToken = rsMaxDBToken!MaxDBToken + 1
'    End If
'
'    rsMaxDBToken.Close
'    Set rsMaxDBToken = Nothing
'
'    sSQL = "INSERT INTO ArezzoToken VALUES ('" & sArezzoId & "'," & lDBToken & "," & lStudyId & ",null,null)"
'    MacroADODBConnection.Execute sSQL
'
'    msArezzoId = sArezzoId
    
Exit Sub
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modDataEntry.InsertArezzoToken"
End Sub

'----------------------------------------------------------------------
Public Sub DeleteArezzoToken()
'----------------------------------------------------------------------
' REM 24/04/02
' Delete our current Arezzo token
' NCJ 22 Aug 06 - Use LOCK DLL to do this; reset ArezzoID and StudyID
'----------------------------------------------------------------------
'Dim sSQL As String

    On Error GoTo Errorlabel

    If msArezzoId <> "" Then
        'only delete if not blank
        Call MACROLOCKBS30.CacheRemoveSubjectRow(goUser.CurrentDBConString, msArezzoId)
        msArezzoId = ""
        mlStudyId = 0
        
'        sSQL = "DELETE FROM ArezzoToken " _
'            & " WHERE "
'        If lClinicalTrialId <> 0 Then
'            'if clincialtrialid named then use it
'            sSQL = sSQL & "ClinicalTrialId = " & lClinicalTrialId & " AND "
'        End If
'        sSQL = sSQL & "ArezzoID = '" & msArezzoId & "'"
'
'        MacroADODBConnection.Execute sSQL
        
    End If

Exit Sub
  
Errorlabel:
    If MACROErrorHandler("modDataEntry", Err.Number, Err.Description, "DeleteArezzoToken", Err.Source) = Retry Then
        Resume
    End If

End Sub

'----------------------------------------------------------------------------------------'
Private Sub StoreStudyIDinReg(lStudyId As Long)
'----------------------------------------------------------------------------------------'
'REM 24/04/02
'Saves the Study Id to the registry
'----------------------------------------------------------------------------------------'
 Dim sSetting As String
    
    On Error GoTo Errorlabel
        
    sSetting = CStr(lStudyId)
    
    Call goUser.UserSettings.SetSetting(SETTING_LAST_USED_STUDY, lStudyId)
    
Exit Sub
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modDataEntry.StoreStudyIDinReg"
End Sub

'----------------------------------------------------------------------------------------'
Private Function GetStudyIDfromReg() As Long
'----------------------------------------------------------------------------------------'
'REM 24/04/02
'Get the StudyId from the registry.  The Study Id in the registry is the last one used.
'----------------------------------------------------------------------------------------'

    On Error GoTo Errorlabel

    GetStudyIDfromReg = CLng(goUser.UserSettings.GetSetting(SETTING_LAST_USED_STUDY, 0))
    
Exit Function
Errorlabel:
    'any problems? - give back zero
    GetStudyIDfromReg = 0
    
End Function

'----------------------------------------------------------------------
Public Function CheckStudyDefCurrent(lStudyId As Long) As Boolean
'----------------------------------------------------------------------
' REM 2/04/02
' Check to see if the StudyDef is not out of date
' NCJ 22 Aug 06 - Use the LOCK DLL to do this; made Public so can call from modSchedule
'----------------------------------------------------------------------
'Dim sSQL As String
'Dim rsCount As ADODB.Recordset
'Dim nCount As Integer

    On Error GoTo Errorlabel
    
    ' NCJ 22 Aug 06 - Check we're talking about the same study!
    If lStudyId <> mlStudyId Then
        CheckStudyDefCurrent = False
    Else
        CheckStudyDefCurrent = MACROLOCKBS30.CacheEntryStillValid(goUser.CurrentDBConString, msArezzoId)
    End If
    
'    'Count all the rows that contain the passed in Study Id and arezzoId
'    sSQL = "SELECT COUNT(*) FROM ArezzoToken" _
'        & " WHERE ClinicalTrialId = " & lStudyId _
'        & " AND ArezzoId = '" & msArezzoId & "'"
'
'    Set rsCount = New ADODB.Recordset
'    rsCount.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    nCount = rsCount.Fields(0).Value
'
'    rsCount.Close
'    Set rsCount = Nothing
'
'    Select Case nCount
'    Case 0
'        'if count is 0 then the studydef is out of date
'        CheckStudyDefCurrent = False
'    Case 1
'        'if the count is one the studydef is current
'        CheckStudyDefCurrent = True
'    Case Is > 1
'        'should never happen, but if it does clear up the table and force a refresh
'        sSQL = "DELETE FROM ArezzoToken" _
'        & " WHERE ClinicalTrialId = " & lStudyId _
'        & " AND ArezzoId = '" & msArezzoId & "'"
'        MacroADODBConnection.Execute sSQL
'
'        CheckStudyDefCurrent = False
'    End Select
    
Exit Function
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modDataEntry.CheckStudyDefCurrent"
End Function

'----------------------------------------------------------------------
Public Function LoadStudy(oStudyDef As StudyDefRO, sCon As String, _
                          lStudyId As Long, nVersionId As Integer, oArezzo As Arezzo_DM, _
                          bGetIdFromReg As Boolean) As String
'----------------------------------------------------------------------
'TA CBB 2.2.20.3 18/7/02: only insert a row into the arezzo table if we load a study
'----------------------------------------------------------------------
    On Error GoTo Errorlabel
    
    If bGetIdFromReg Then
        lStudyId = GetStudyIDfromReg
        'check study Id exists
        If TrialNameFromId(lStudyId) = "" Then
            lStudyId = 0
        End If
    End If
    
    If lStudyId <> 0 Then
        'NB we use a study id of 0 to indicate that we will not load the study
        'this will only happen when called from frmMenu.InitialiseMe
        If oStudyDef Is Nothing Then
            Set oStudyDef = New StudyDefRO
        End If
        LoadStudy = oStudyDef.Load(sCon, lStudyId, nVersionId, oArezzo)
        'TA CBB 2.2.20.3 18/7/02: only insert a row into the arezzo table if we load a study
#If VTRACK <> 1 Then
        Call InsertArezzoToken(lStudyId)
#End If
    End If
    

Exit Function
Errorlabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modDataEntry.LoadStudy"
End Function

' NCJ 22 Aug 06 - Token not needed any more
''----------------------------------------------------------------------
'Private Function Token() As String
''----------------------------------------------------------------------
'' Calculate a new token
''----------------------------------------------------------------------
'Dim nAscStart As Integer
'Dim i As Long
'Dim sToken As String
'
'    'ensure that a random token is generated
'    Randomize
'
'    For i = 1 To 10
'        sToken = sToken & Chr$(65 + RndLong(25, True))
'    Next
'
'    Token = sToken
'
'End Function
'

'---------------------------------------------------------------------
Public Function LoadSubject(ByRef oStudyDef As StudyDefRO, ByRef oArezzo As Arezzo_DM, _
                                ByRef sSubjectToken As String, _
                                lStudyId As Long, sSite As String, _
                                Optional lSubjectId As Long = g_NEW_SUBJECT_ID) As Boolean
'---------------------------------------------------------------------
'todo add readonly parameter to this function for the moment to keep compatabiltiy i have made this optional
' this should not neccesarily be optional in final version
'new hourglass form displayed
' NCJ 24 Jan 03 - Added sCountry to NewSubject
' NCJ 6 May 03 - Added new parameters sUserNameFull and sUserRole to Load/New subject
' NCJ 22 Aug 06 - Don't need StudyID for DeleteArezzoToken
'---------------------------------------------------------------------
Dim sLockErrMsg As String
Dim sErrMsg As String
Dim enUpdateMode As eUIUpdateMode
Dim sCountry As String
'version id currently always 1
Const nVERSION_ID As Integer = 1

    On Error GoTo ErrLabel

    If CheckStudyDefCurrent(lStudyId) Then
         'TA 22/11/2001: use new routine to unload subject but NOT study
        If Not CloseSubject(oStudyDef, False, False, True) Then
            'this should never happen
            Err.Raise vbObjectError + 1002, , "load subject routine could not close subject"
        End If

    Else
         'TA 22/11/2001: use new routine to unload subject AND study
        If Not CloseSubject(oStudyDef, False, True, True) Then
            'this should never happen
            Err.Raise vbObjectError + 1002, , "load subject routine could not close subject"
        End If
        'load study
        Set oStudyDef = New StudyDefRO
        frmHourglass.Display "Loading study definition", True
        sErrMsg = LoadStudy(oStudyDef, gsADOConnectString, lStudyId, nVERSION_ID, oArezzo, False)
        UnloadfrmHourglass
        If sErrMsg <> "" Then
    #If VTRACK <> 1 Then
            ' NCJ 22 Aug 06 - Don't need StudyID for DeleteArezzoToken
            Call DeleteArezzoToken
    #End If
            DialogError sErrMsg
            oStudyDef.Terminate
            Set oStudyDef = Nothing
            LoadSubject = False
            Exit Function
        End If
    End If

    ' NCJ 22 Mar 02 - Set the update mode
    If goUser.CheckPermission(gsFnChangeData) _
            And Not (goUser.DBIsServer And goUser.GetAllSites.Item(sSite).SiteLocation = 1) Then
            'change data and not (is server and site is remote)
        enUpdateMode = eUIUpdateMode.Read_Write
    Else
        enUpdateMode = eUIUpdateMode.Read_Only
    End If
    
    If lSubjectId = g_NEW_SUBJECT_ID Then
        ' NCJ 24 Jan 03 - Added sCountry
        sCountry = goUser.GetAllSites.Item(sSite).CountryName
        Call oStudyDef.NewSubject(sSite, goUser.UserName, sCountry, goUser.UserNameFull, goUser.UserRole)
    Else
        frmHourglass.Display "Loading subject", True
        Call oStudyDef.LoadSubject(sSite, lSubjectId, goUser.UserName, enUpdateMode, goUser.UserNameFull, goUser.UserRole)
        UnloadfrmHourglass
    End If
    
    If oStudyDef.Subject.CouldNotLoad Then
        DialogWarning oStudyDef.Subject.CouldNotLoadReason
        'TA 03/04/2003: remove subject from memory so that IsSubjectOpen works
        oStudyDef.RemoveSubject
        LoadSubject = False
    Else
        ' Assume there is a valid subject at this point!
        ' Initialise registration
        ' NB This also enables/disables the "Register subject" menu item
        Call InitialiseRegistration(oStudyDef.Subject)
        'for the moment a lock error means the subject failed to load r/w
        LoadSubject = True
        
        ' NCJ 23 Jan 03 - Only set LocalDateFormat if user wants to use it
        If goUser.UserSettings.GetSetting(SETTING_LOCAL_FORMAT, False) Then
            oStudyDef.Subject.LocalDateFormat = goUser.UserSettings.GetSetting(SETTING_LOCAL_DATE_FORMAT, "")
        End If
        
        ' NCJ 24 Jan 03 - Add current user details
        ' NCJ 6 May 03 - Don't need this as now passed as params to New/Load subject
'        Call oStudyDef.Subject.SetUserProperties(goUser.UserNameFull, goUser.UserRole)

    End If
    
    Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modDataEntry.LoadSubject"
    
End Function

'---------------------------------------------------------------------
Public Function QuestionNameText(oResponse As Response) As String
'---------------------------------------------------------------------
' Get the display text for this question name
'---------------------------------------------------------------------
Dim sText As String

    sText = oResponse.Element.Name
    ' NCJ 15 Aug 02 - For group questions, show repeat number
    If Not oResponse.Element.OwnerQGroup Is Nothing Then
        sText = sText & " [" & oResponse.RepeatNumber & "]"
    End If

    QuestionNameText = sText

End Function

'--------------------------------------------------------------------------------------------------
Public Function RtnSubjectText(ByVal sSubjectId As String, ByVal vSubjectLabel As Variant) As String
'--------------------------------------------------------------------------------------------------
' ic 14/02/2003 function returns a subject label or "(" id ")"
'--------------------------------------------------------------------------------------------------
Dim sRtn As String

    If Not IsNull(vSubjectLabel) Then
        If (vSubjectLabel <> "") Then sRtn = vSubjectLabel
    End If
    If (sRtn = "") Then
        sRtn = "(" & sSubjectId & ")"
    End If
    RtnSubjectText = sRtn
End Function
