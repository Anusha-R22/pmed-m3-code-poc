Attribute VB_Name = "modSchedule"
'----------------------------------------------------
' File:         modSchedule.bas (used to be frmSchedule)
' Copyright:    InferMed 2000-2003, All Rights Reserved
' Author:       Nicky Johns/Toby Aldridge, InferMed, May 2001
'               Form to display Subject Schedule for MACRO 3.0
'----------------------------------------------------
' REVISIONS
'   NCJ 30/8/01 - Some tidying up
'   DPH 12/10/2001 - CanOpenEform & call to it in flxSchedule_DblClick to
'       control opening new forms
'   DPH 24/10/2001 - flxSchedule_MouseMove edited to correct tooltip problem
'   NCJ 20 Mar 02 - Ensure visit/form dates only editable when appropriate
'   NCJ 15 May 02 - Conditionally compiled code for Multi User support
'   za 22/05/02, only call RedrawGrid if there is a Subject
'   TA 11/07/02: CBB 2.2.19.12 Allow changing of status by right-clicking an eForm in the schedule
'   NCJ 14 Aug 02 - Changed ToggleEFIStatus to use new Response status routine
'   MLM 30/08/02: Don't check for a visit date in flxSchedule_DblClick()
'   NCJ 18 Sept 02 - frmEFormDataEntry.Display may fail if we failed to get Responses
'   NCJ 20 Sept 02 - Added extra arg. to RemoveResponses
'   TA 26/09/02: Changes for New UI - no title bar, not maximised etc
'   NCJ 26 Sept 02 - Removed EditDate & associated stuff; added check for subject updates when double clicking
'   NCJ 30 Sept 02 - Fixed bug in ScheduleUpdatedByOtherUser
'   TA 01/10/2002 - Converted from a form to a module
'   TA 04/10/2002: Change to use new clsMenuItem for PopUpMenus
' TA 15/10/2002: Mores changes to the User Interface
'   NCJ 16 Oct 02 - Added extra menu items to test SDVs for eForms, Visit, Subjects
'   NCJ 5 Nov 02 - Do not allow more than one SDV per object
'   TA 03/01/2003: New routines to display right click subject and visit menus
'   DPH 14/01/2003 Lock/Unlock/Freeze menu items
'   DPH 21/01/2003 Unobtainable at Subject/Visit Level + requested to unobtainable
'   NCJ 22 Jan 03 - Make sure we don't allow users to open empty locked or frozen eForms
'               Change message when there are no objects to be made Unobtainable/Missing
'   NCJ 23 Jan 03 - Ensure VisitEForm is processed first when making Visit Unobtainable
'   NCJ 10 Feb 03 - Made CanOpenEForm Public, and added User & Subject parameters
' TA 03/04/2003 - No need to unload webbrowser forms any more before refreshing them
'   NCJ 12 Aug 04 - Added right mouse menu item to change all Planned question SDVS to Done for eForm
'   NCJ 24 Aug 04 - Changed gsFnLockData to gsFnUnLockData in all the "UNLOCK" menu items (Bug 2368)
'   NCJ 22 Aug 06 - Deal properly with changes to the StudyDef (Bug 2786)
'   NCJ 31 Aug 07 - Issue 2938 - Added "Revalidate subject" to subject right-mouse menu
'----------------------------------------------------

Option Explicit

Private moSubject As StudySubject
Private moUser As MACROUser

Private mofrmSchedule As frmWebBrowser
Private mofrmeFormTop As frmWebBrowser
Private mofrmeFormLh As frmWebBrowser

'----------------------------------------------------------------
Public Function ShowSchedule(oUser As MACROUser, oStudyDef As StudyDefRO, ofrmSchedule As frmWebBrowser, _
                                        ofrmeFormTop As frmWebBrowser, ofrmeFormLh As frmWebBrowser)
'----------------------------------------------------------------
' Display an schedule for a preloaded subject
'----------------------------------------------------------------

       
    'create  a reference to the subject
    Set moSubject = oStudyDef.Subject
    Set moUser = oUser
    Set mofrmSchedule = ofrmSchedule
    Set mofrmeFormTop = ofrmeFormTop
    Set mofrmeFormLh = ofrmeFormLh
 

    RefreshSchedule

    
End Function

'----------------------------------------------------------------
Public Function CloseSchedule()
'----------------------------------------------------------------
' Close schedule
'----------------------------------------------------------------
       
    'tidy up refenences
    
    If Not mofrmSchedule Is Nothing Then
        'TA blank out schedule for speed
        mofrmSchedule.Display wdtHTML, "", "no"
        mofrmSchedule.Visible = False
    End If
    Set moSubject = Nothing
    Set moUser = Nothing
    Set mofrmSchedule = Nothing
    Set mofrmeFormTop = Nothing
    Set mofrmeFormLh = Nothing
    
    
End Function

'--------------------------------------------
Public Function RefreshSchedule() As Boolean
'--------------------------------------------
' Refresh the schedule grid display.
' ReUse RedrewGrid if the form is visible
'--------------------------------------------
Dim sInnerScheduleHTML As String
Dim lTop As Long
Dim lWidth As Long
Dim lLeft As Long
Dim lHeight As Long
Dim sScheduleHTML As String

    On Error GoTo ErrLabel
    
    If ScheduleOpen Then
        
        HourglassOn

        'get schedule html before unload - for speed
        sScheduleHTML = ScheduleHTML(moSubject)
        
        Call mofrmSchedule.Display(wdtHTML, sScheduleHTML, "auto")
        'MLM 12/02/03: After loading the HTML, call javascript to expand markup
        mofrmSchedule.ExecuteJavaScript.fnPageLoaded
        
        HourglassOff

    End If
    
    RefreshSchedule = ScheduleOpen
    
    Exit Function
    
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modSchedule.RefreshSchedule"

End Function

'----------------------------------------------------------------------------------------------
Private Property Get ScheduleOpen() As Boolean
'----------------------------------------------------------------------------------------------
'is shceudle open
'----------------------------------------------------------------------------------------------

    ScheduleOpen = Not (mofrmSchedule Is Nothing)
    
End Property

'----------------------------------------------------------------------------------------------
Public Sub DisplayScheduleMenu(lEFITaskId As Long)
'----------------------------------------------------------------------------------------------
'Display pop up menu for eForms
'   TA 11/07/02: CBB 2.2.19.12 Allow changing of status by right-clicking an eForm in the schedule
' DPH 14/01/2003 Lock/Unlock/Freeze menu items
' NCJ 12 Aug 04 - Change Planned SDVs to Done
'----------------------------------------------------------------------------------------------
Dim oEFI As EFormInstance
Dim sMenuItemSelected As String
Dim oMenuItem As clsMenuItem
Dim oMenuItems As clsMenuItems

Dim eLockFreezeStatus As eLockStatus
Dim bLocked As Boolean
Dim bUnLocked As Boolean
Dim bFrozen As Boolean
Dim bCanUnfreeze As Boolean
Dim lVisitTaskId As Long
Dim oVI As VisitInstance

    On Error GoTo ErrLabel
   
    Set oEFI = moSubject.eFIByTaskId(lEFITaskId)
        
    ' It is an eForm
    If Not oEFI Is Nothing Then
        ' See if we have a locked or frozen item
        bLocked = False
        bFrozen = False
        eLockFreezeStatus = oEFI.LockStatus
        Select Case eLockFreezeStatus
            Case eLockStatus.lsLocked
                bLocked = True
            Case eLockStatus.lsFrozen
                bFrozen = True
            Case Else
        End Select
        bUnLocked = Not bLocked And Not bFrozen
        lVisitTaskId = oEFI.VisitInstance.VisitTaskId
        Set oVI = oEFI.VisitInstance
        
        ' NCJ 23 Dec 02 - Can only unfreeze if it's the subject OR our parent is not frozen
        ' (User's Unfreeze permission is checked later)
        If bFrozen Then
            ' unfreeze if subject/visit not frozen
            bCanUnfreeze = ((moSubject.LockStatus <> eLockStatus.lsFrozen) And _
                            (oVI.LockStatus <> eLockStatus.lsFrozen))
        Else
            ' Can't unfreeze if not frozen!
            bCanUnfreeze = False
        End If
        
        Set oMenuItems = New clsMenuItems
        Call oMenuItems.Add("OPEN", "Open...", CanOpenEform(oEFI, moSubject, moUser, ""), False, True)
        Call oMenuItems.AddSeparator
        Call oMenuItems.Add("LOCK", "Lock", _
                    goUser.CheckPermission(gsFnLockData) And bUnLocked)
        ' NCJ 12 Aug 04 - Changed gsFnLockData to gsFnUnLockData
        Call oMenuItems.Add("UNLOCK", "Unlock", _
                    goUser.CheckPermission(gsFnUnLockData) And bLocked)
        Call oMenuItems.Add("FREEZE", "Freeze", _
                    goUser.CheckPermission(gsFnFreezeData) And Not bFrozen)
        Call oMenuItems.Add("UNFREEZE", "Unfreeze", _
                    goUser.CheckPermission(gsFnUnFreezeData) And bCanUnfreeze)
        Call oMenuItems.AddSeparator
        Call oMenuItems.Add("UNOB", "Unobtainable", CanMakeEFormUnobtainable(oEFI), False)
        Call oMenuItems.Add("MISS", "Missing", CanMakeEFormMissing(oEFI), False)
        ' NCJ 16 Oct 02 - These SDV items are here temporarily and will probably end up somewhere else...
        Call oMenuItems.AddSeparator
        ' MLM 29/06/05: bug 2464: Show these menu items conditionally depending on a new MACRO setting
        '   and add the '/Edit' wording to menu items
        Call oMenuItems.Add("SDV3", "Create/Edit eForm SDV...", goUser.CheckPermission(gsFnCreateSDV), False)
        If (LCase(GetMACROSetting(MACRO_SETTING_SHOW_SDV_SCHEDULE_MENU, "true")) = "true") Then
            Call oMenuItems.Add("SDV2", "Create/Edit Visit SDV...", goUser.CheckPermission(gsFnCreateSDV), False)
            Call oMenuItems.Add("SDV1", "Create/Edit Subject SDV...", goUser.CheckPermission(gsFnCreateSDV), False)
        End If
        ' NCJ 12 Aug 04 - Change Planned SDVs to Done
        Call oMenuItems.AddSeparator
        Call oMenuItems.Add("QSDVS", "Change all Planned question SDVs to Done...", _
                        goUser.CheckPermission(gsFnCreateSDV) And CanChangePlannedSDVs(oEFI), False)
    
        ' Show popup mennu
        sMenuItemSelected = frmMenu.ShowPopUpMenu(oMenuItems)
        ' Refesh form (so that bit ofform under popup menu reshown
        mofrmSchedule.Refresh
        Select Case sMenuItemSelected
        Case "OPEN" 'open
            Call ScheduleOpeneForm(lEFITaskId)
        Case "UNOB" 'set eform to unobtainable
            Call ToggleEFIStatus(oEFI, eStatus.Unobtainable, True, True)
        Case "MISS" 'set eform to missing
            Call ToggleEFIStatus(oEFI, eStatus.Missing, True, True)
            
        Case "SDV3" 'Create SDV for eForm
            Call NewEFormMIMessage(oEFI, MIMsgType.mimtSDVMark)
        Case "SDV2" 'Create SDV for Visit
            Call NewVisitMIMessage(oEFI.VisitInstance, MIMsgType.mimtSDVMark)
        Case "SDV1" 'Create SDV for Subject
            Call NewSubjectMIMessage(moSubject, MIMsgType.mimtSDVMark)
        Case "QSDVS"    ' NCJ 12 Aug 04 - Change Planned SDVs to Done
            Call ChangeSDVsToDone(oEFI)
        Case "LOCK" ' NCJ 23 Dec 02 - Pass new LFAction rather than LockStatus
            Call ChangeLockUnlockFreeze(LFAction.lfaLock, lfscEForm, moSubject.StudyId, _
                    moSubject.StudyCode, moSubject.Site, moSubject.PersonId, _
                    oVI.VisitId, _
                    oVI.CycleNo, _
                    oEFI.EFormTaskId, oEFI.eForm.EFormId, oEFI.CycleNo)
        Case "UNLOCK"
            Call ChangeLockUnlockFreeze(LFAction.lfaUnlock, lfscEForm, moSubject.StudyId, _
                    moSubject.StudyCode, moSubject.Site, moSubject.PersonId, _
                    oVI.VisitId, _
                    oVI.CycleNo, _
                    oEFI.EFormTaskId, oEFI.eForm.EFormId, oEFI.CycleNo)
        Case "FREEZE"
            Call ChangeLockUnlockFreeze(LFAction.lfaFreeze, lfscEForm, moSubject.StudyId, _
                    moSubject.StudyCode, moSubject.Site, moSubject.PersonId, _
                    oVI.VisitId, _
                    oVI.CycleNo, _
                    oEFI.EFormTaskId, oEFI.eForm.EFormId, oEFI.CycleNo)
        Case "UNFREEZE"
            Call ChangeLockUnlockFreeze(LFAction.lfaUnfreeze, lfscEForm, moSubject.StudyId, _
                    moSubject.StudyCode, moSubject.Site, moSubject.PersonId, _
                    oVI.VisitId, _
                    oVI.CycleNo, _
                    oEFI.EFormTaskId, oEFI.eForm.EFormId, oEFI.CycleNo)
        End Select
        
    End If

    Exit Sub
    
ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "modSchedule.DisplayScheduleMenu"
    
End Sub

'----------------------------------------------------------------------------------------------
Public Sub DisplayVisitMenu(lVisitTaskId As Long)
'----------------------------------------------------------------------------------------------
'TA 03/01/2003: Display pop up menu for editing visit data
'DPH 14/01/2003 Lock/Unlock/Freeze menu items
'----------------------------------------------------------------------------------------------
Dim sMenuItemSelected As String
Dim oMenuItem As clsMenuItem
Dim oMenuItems As clsMenuItems

Dim oVI As VisitInstance
Dim eLockFreezeStatus As eLockStatus
Dim bLocked As Boolean
Dim bUnLocked As Boolean
Dim bFrozen As Boolean
Dim bCanUnfreeze As Boolean
    
    On Error GoTo ErrLabel
   
    ' See if we have a locked or frozen item
    bLocked = False
    bFrozen = False
    Set oVI = moSubject.VisitInstanceByTaskId(lVisitTaskId)
    'eLockFreezeStatus = moSubject.VisitInstanceByTaskId(lVisitTaskId).LockStatus
    eLockFreezeStatus = oVI.LockStatus
    Select Case eLockFreezeStatus
        Case eLockStatus.lsLocked
            bLocked = True
        Case eLockStatus.lsFrozen
            bFrozen = True
        Case Else
    End Select
    bUnLocked = Not bLocked And Not bFrozen

    ' NCJ 23 Dec 02 - Can only unfreeze if it's the subject OR our parent is not frozen
    ' (User's Unfreeze permission is checked later)
    If bFrozen Then
        ' unfreeze if subject not frozen
        bCanUnfreeze = (moSubject.LockStatus <> eLockStatus.lsFrozen)
    Else
        ' Can't unfreeze if not frozen!
        bCanUnfreeze = False
    End If
    
    Set oMenuItems = New clsMenuItems
    Call oMenuItems.Add("LOCK", "Lock", _
                goUser.CheckPermission(gsFnLockData) And bUnLocked)
    ' NCJ 24 Aug 04 - Changed gsFnLockData to gsFnUnLockData
    Call oMenuItems.Add("UNLOCK", "Unlock", _
                goUser.CheckPermission(gsFnUnLockData) And bLocked)
    Call oMenuItems.Add("FREEZE", "Freeze", _
                goUser.CheckPermission(gsFnFreezeData) And Not bFrozen)
    Call oMenuItems.Add("UNFREEZE", "Unfreeze", _
                goUser.CheckPermission(gsFnUnFreezeData) And bCanUnfreeze)
    Call oMenuItems.AddSeparator
    Call oMenuItems.Add("UNOB", "Unobtainable", CanMakeVisitUnobtainable(oVI), False)
    Call oMenuItems.Add("MISS", "Missing", CanMakeVisitMissing(oVI), False)
    Call oMenuItems.AddSeparator
    ' MLM 01/07/05: Changed text to reflect that edit is now possible:
    Call oMenuItems.Add("SDV", "Create/Edit visit SDV...", goUser.CheckPermission(gsFnCreateSDV), False)

    'show popup mennu
    sMenuItemSelected = frmMenu.ShowPopUpMenu(oMenuItems)
    'refesh form (so that bit ofform under popup menu reshown
    mofrmSchedule.Refresh
    Select Case sMenuItemSelected
    Case "SDV" 'Create SDV for Visit
        Call NewVisitMIMessage(oVI, MIMsgType.mimtSDVMark)
    Case "LOCK" ' NCJ 23 Dec 02 - Pass new LFAction rather than LockStatus
        Call ChangeLockUnlockFreeze(LFAction.lfaLock, lfscVisit, moSubject.StudyId, _
                moSubject.StudyCode, moSubject.Site, moSubject.PersonId, _
                oVI.VisitId, _
                oVI.CycleNo)
    Case "UNLOCK"
        Call ChangeLockUnlockFreeze(LFAction.lfaUnlock, lfscVisit, moSubject.StudyId, _
                moSubject.StudyCode, moSubject.Site, moSubject.PersonId, _
                oVI.VisitId, _
                oVI.CycleNo)
    Case "FREEZE"
        Call ChangeLockUnlockFreeze(LFAction.lfaFreeze, lfscVisit, moSubject.StudyId, _
                moSubject.StudyCode, moSubject.Site, moSubject.PersonId, _
                oVI.VisitId, _
                oVI.CycleNo)
    Case "UNFREEZE"
        Call ChangeLockUnlockFreeze(LFAction.lfaUnfreeze, lfscVisit, moSubject.StudyId, _
                moSubject.StudyCode, moSubject.Site, moSubject.PersonId, _
                oVI.VisitId, _
                oVI.CycleNo)
    Case "UNOB" 'set visit to unobtainable
        Call ToggleVisitStatus(oVI, eStatus.Unobtainable)
    Case "MISS" 'set visit to missing
        Call ToggleVisitStatus(oVI, eStatus.Missing)
    End Select


    Exit Sub
    
ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "modSchedule.DisplayScheduleMenu"
    
End Sub


'----------------------------------------------------------------------------------------------
Public Sub DisplaySubjectMenu()
'----------------------------------------------------------------------------------------------
'TA 03/01/2003: Display pop up menu for editing subject data
'DPH 14/01/2003 Lock/Unlock/Freeze menu items
' NCJ 31 Aug 07 - Bug 2938 - Added Revalidate subject
'----------------------------------------------------------------------------------------------
Dim sMenuItemSelected As String
Dim oMenuItem As clsMenuItem
Dim oMenuItems As clsMenuItems

Dim eLockFreezeStatus As eLockStatus
Dim bLocked As Boolean
Dim bUnLocked As Boolean
Dim bFrozen As Boolean
Dim bCanUnfreeze As Boolean
    
    On Error GoTo ErrLabel
    
    ' See if we have a locked or frozen item
    bLocked = False
    bFrozen = False
    eLockFreezeStatus = moSubject.LockStatus
    Select Case eLockFreezeStatus
        Case eLockStatus.lsLocked
            bLocked = True
        Case eLockStatus.lsFrozen
            bFrozen = True
        Case Else
    End Select
    bUnLocked = Not bLocked And Not bFrozen

    ' NCJ 23 Dec 02 - Can only unfreeze if it's the subject OR our parent is not frozen
    ' (User's Unfreeze permission is checked later)
    If bFrozen Then
        ' We can unfreeze the subject
        bCanUnfreeze = True
    Else
        ' Can't unfreeze if not frozen!
        bCanUnfreeze = False
    End If

    Set oMenuItems = New clsMenuItems
    Call oMenuItems.Add("LOCK", "Lock", _
                goUser.CheckPermission(gsFnLockData) And bUnLocked)
    ' NCJ 24 Aug 04 - Changed gsFnLockData to gsFnUnLockData
    Call oMenuItems.Add("UNLOCK", "Unlock", _
                goUser.CheckPermission(gsFnUnLockData) And bLocked)
    Call oMenuItems.Add("FREEZE", "Freeze", _
                goUser.CheckPermission(gsFnFreezeData) And Not bFrozen)
    Call oMenuItems.Add("UNFREEZE", "Unfreeze", _
                goUser.CheckPermission(gsFnUnFreezeData) And bCanUnfreeze)
    Call oMenuItems.AddSeparator
    Call oMenuItems.Add("UNOB", "Unobtainable", CanMakeSubjectUnobtainable(), False)
    Call oMenuItems.Add("MISS", "Missing", CanMakeSubjectMissing(), False)
    Call oMenuItems.AddSeparator
    ' MLM 01/07/05: Changed text to reflect that edit is now possible:
    Call oMenuItems.Add("SDV", "Create/Edit subject SDV...", goUser.CheckPermission(gsFnCreateSDV) And Not bFrozen, False)
    ' NCJ 31 Aug 07 - 2938 - Added Revalidate
    Call oMenuItems.AddSeparator
    Call oMenuItems.Add("REVAL", "Revalidate subject...", CanRevalidateSubject(), False)
    
    'show popup mennu
    sMenuItemSelected = frmMenu.ShowPopUpMenu(oMenuItems)
    'refesh form (so that bit ofform under popup menu reshown
    mofrmSchedule.Refresh
    Select Case sMenuItemSelected
    Case "SDV" 'Create SDV for Subject
        Call NewSubjectMIMessage(moSubject, MIMsgType.mimtSDVMark)

    Case "LOCK" ' NCJ 23 Dec 02 - Pass new LFAction rather than LockStatus
        Call ChangeLockUnlockFreeze(LFAction.lfaLock, lfscSubject, moSubject.StudyId, _
                moSubject.StudyCode, moSubject.Site, moSubject.PersonId)
    Case "UNLOCK"
        Call ChangeLockUnlockFreeze(LFAction.lfaUnlock, lfscSubject, moSubject.StudyId, _
                moSubject.StudyCode, moSubject.Site, moSubject.PersonId)
    Case "FREEZE"
        Call ChangeLockUnlockFreeze(LFAction.lfaFreeze, lfscSubject, moSubject.StudyId, _
                moSubject.StudyCode, moSubject.Site, moSubject.PersonId)
    Case "UNFREEZE"
        Call ChangeLockUnlockFreeze(LFAction.lfaUnfreeze, lfscSubject, moSubject.StudyId, _
                moSubject.StudyCode, moSubject.Site, moSubject.PersonId)
    Case "UNOB" 'set subject to unobtainable
        Call ToggleSubjectStatus(moSubject, eStatus.Unobtainable)
    Case "MISS" 'set subject to missing
        Call ToggleSubjectStatus(moSubject, eStatus.Missing)
    Case "REVAL" ' Revalidate this subject
        ' Check they want to do it
        If DialogQuestion("All eForms for this subject will be revalidated, and all changes will be saved." _
                & vbCrLf & "Are you sure you wish to continue?") = vbYes Then
            Call RevalidateSubject(moSubject)
        End If
    End Select


    Exit Sub
    
ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|" & "modSchedule.DisplaySubjectMenu"
    
End Sub

'--------------------------------------------------------------
Private Function CanRevalidateSubject() As Boolean
'--------------------------------------------------------------
' NCJ 31 Aug 07 - 2938 - Can we revalidate this subject?
'--------------------------------------------------------------

    CanRevalidateSubject = False
    
    ' User must be able to change data
    If Not goUser.CheckPermission(gsFnChangeData) Then Exit Function
    
    ' Subject must not be Read-Only
    If moSubject.ReadOnly Then Exit Function
    
    ' Subject must not be locked or frozen
    If moSubject.LockStatus <> eLockStatus.lsUnlocked Then Exit Function
    
    ' If we get here, all is OK
    CanRevalidateSubject = True

End Function

'--------------------------------------------------------------
Private Sub RevalidateSubject(oSubject As StudySubject)
'--------------------------------------------------------------
' NCJ 31 Aug 07 - Revalidate this subject
'--------------------------------------------------------------
Dim oRevalidator As Revalidator
Dim sLogFile As String
Dim sViewMsg As String

    On Error GoTo ErrLabel
    
    HourglassOn
    
    frmHourglass.Display "Revalidating subject...", True
    ' Don't let them click anywhere and mess things up
    Call frmMenu.ToggleUserIntervention(False)
    
    ' Select a suitable log file name
    sLogFile = gsTEMP_PATH & oSubject.StudyCode & "_" & oSubject.Site & "_" & oSubject.PersonId _
                    & "_" & Format(Now, "yyyymmddhhmmss") & ".txt"
    
    Set oRevalidator = New Revalidator
    ' Default to non-verbose, i.e. changes only
    oRevalidator.Verbose = False
    Call oRevalidator.InitRevalidation(sLogFile, goUser, "")
    
    Call oRevalidator.Revalidate(goUser, oSubject.StudyId, oSubject.Site, oSubject.PersonId, _
                        oSubject.StudyCode, oSubject.label, goArezzo, 0)
    Call oRevalidator.EndRevalidation
    
    HourglassOff
    
    UnloadfrmHourglass
    Call frmMenu.ToggleUserIntervention(True)

    
    ' Show the log file
    sViewMsg = "Revalidation report has been saved to the MACRO 'Temp' folder." & vbCrLf
    sViewMsg = sViewMsg & "Would you like to view the report?"
    
    If DialogQuestion(sViewMsg) = vbYes Then
        ' Open document in Notepad (or default text editor)
        Call ShowDocument(frmMenu.hWnd, sLogFile)
    End If
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "modSchedule.RevalidateSubject"

End Sub
       
'--------------------------------------------------------------
Public Function ScheduleOpeneForm(lEFITaskId As Long) As Boolean
'--------------------------------------------------------------
' Opens the eForm with the supplied eFormID
' if it hasn't been made out of date
' NCJ 22 Jan 03 - Added check that eForm CAN be opened
'--------------------------------------------------------------
Dim oEFI As EFormInstance
Dim sMsg As String

    ' If the schedule's changed we have to start again
    If Not ScheduleUpdatedByOtherUser("Please reselect the eForm to be opened.") Then
        Set oEFI = moSubject.eFIByTaskId(lEFITaskId)
        If Not CanOpenEform(oEFI, moSubject, moUser, sMsg) Then
            DialogInformation sMsg
            ScheduleOpeneForm = False
        Else
            With mofrmSchedule
                ScheduleOpeneForm = frmEFormDataEntry.Display(moUser, oEFI, _
                    .Left + WEB_EFORM_LH_WIDTH, .Top + WEB_EFORM_TOP_HEIGHT, _
                    .Width - WEB_EFORM_LH_WIDTH, .Height - WEB_EFORM_TOP_HEIGHT, _
                     mofrmeFormTop, mofrmeFormLh)
                
            End With
        End If
    Else
        ScheduleOpeneForm = False
    End If

    Set oEFI = Nothing

End Function

'--------------------------------------------------------------
Private Function CanMakeEFormUnobtainable(oEFI As EFormInstance) As Boolean
'--------------------------------------------------------------
' Returns TRUE if current user can unobtainablise the given eForm instance
'--------------------------------------------------------------
' REVISIONS
' DPH 21/01/2003 - Allow Requested form to unobtainable
'--------------------------------------------------------------
    
    CanMakeEFormUnobtainable = False
    
    ' User must be able to change data
    If Not goUser.CheckPermission(gsFnChangeData) Then Exit Function
    
    ' Subject must not be Read-Only
    If moSubject.ReadOnly Then Exit Function
    
    ' Form must not be locked or frozen
    If oEFI.LockStatus <> eLockStatus.lsUnlocked Then Exit Function
    
    ' Form must have status of missing OR requested
    If ((oEFI.Status <> eStatus.Missing) And (oEFI.Status <> eStatus.Requested)) Then Exit Function
    
    ' If we get here, all is OK
    CanMakeEFormUnobtainable = True

End Function

'--------------------------------------------------------------
Private Function CanMakeEFormMissing(oEFI As EFormInstance) As Boolean
'--------------------------------------------------------------
' Returns TRUE if current user can make missing the given eForm instance
'--------------------------------------------------------------

    CanMakeEFormMissing = False
    
    ' User must be able to change data
    If Not goUser.CheckPermission(gsFnChangeData) Then Exit Function
    
    ' Subject must not be Read-Only
    If moSubject.ReadOnly Then Exit Function
    
    ' Form must not be locked or frozen
    If oEFI.LockStatus <> eLockStatus.lsUnlocked Then Exit Function
    
    ' Form must have status of unobtanable
    If oEFI.Status <> eStatus.Unobtainable Then Exit Function

    ' If we get here, all is OK
    CanMakeEFormMissing = True

End Function

'--------------------------------------------------------------
Private Function CanChangePlannedSDVs(oEFI As EFormInstance) As Boolean
'--------------------------------------------------------------
' Can we change Planned SDVs to Done?
'--------------------------------------------------------------

    CanChangePlannedSDVs = False
    
    ' Must have some responses
    If oEFI.Status = eStatus.Requested Then Exit Function
    
    ' Mustn't have SDV status of None or Done
    If oEFI.SDVStatus = eSDVStatus.ssCancelled Then Exit Function
    If oEFI.SDVStatus = eSDVStatus.ssNone Then Exit Function
    If oEFI.SDVStatus = eSDVStatus.ssComplete Then Exit Function

    ' Otherwise there's a possiblity of some Planned SDVs on the eForm
    CanChangePlannedSDVs = True
    
End Function

'--------------------------------------------------------------
Private Sub ToggleEFIStatus(oEFI As EFormInstance, nToStatus As eStatus, _
                            bShowChangeMessage As Boolean, _
                            bDoVisitEForm As Boolean)
'--------------------------------------------------------------
' Make an eForm unobtainable by changing all its 'Missing' responses to 'Unobtainable'
' or vice versa
' no permissions etc are checked in here as they were checked to enable the menu items
' nb we change derived questions
' NCJ 14 Aug 02 - Use new SetStatusFromSchedule routine of Response object
' DPH 21/01/2003 Additional parameter to avoid pop up question
' NCJ 24 Jan 03 - Added bDoVisitEForm if we want to do the Visit eForm too
'--------------------------------------------------------------
Dim oResponse As Response
Dim lChanged As Long
Dim nOldStatus As eStatus
Dim sStatus As String
Dim sMsg As String
Dim sLockErrMsg As String
Dim sEFILockToken As String
Dim sVEFILockToken As String
Dim bSaveResponses As Boolean
Dim bOldStatusMissingRequested As Boolean

    On Error GoTo ErrLabel
    
    ' NCJ 26 Sept 02 - Check for subject data updates first
    ' and don't continue if there are any
    If ScheduleUpdatedByOtherUser("Please try again") Then
        Exit Sub
    End If
    
    'choose the statuses to look for
    If (nToStatus = eStatus.Missing) Then
        nOldStatus = eStatus.Unobtainable
        bOldStatusMissingRequested = False
    Else
        nOldStatus = eStatus.Missing
        bOldStatusMissingRequested = True
    End If
        
    'load efi's responses
    'we don't need to hold onto the EFILock Token or VEFILockToken
    If moSubject.LoadResponses(oEFI, sLockErrMsg, sEFILockToken, sVEFILockToken) <> lrrReadWrite Then
        If Not bShowChangeMessage Then
            DialogError sLockErrMsg
        End If
'EXIT SUB HERE
        Exit Sub
    End If
    
    lChanged = 0
    ' Loop through each response on the "main" eForm
    For Each oResponse In oEFI.Responses
        If oResponse.LockStatus = eLockStatus.lsUnlocked Then
            If ((bOldStatusMissingRequested) And ((oResponse.Status = eStatus.Missing) Or (oResponse.Status = eStatus.Requested))) _
                Or ((Not bOldStatusMissingRequested) And (oResponse.Status = eStatus.Unobtainable)) Then
                'this response is unlocked so toggle this response's status. nb we change derived questions
                Call oResponse.SetStatusFromSchedule(nToStatus)
                'increase the count of responses changed
                lChanged = lChanged + 1
            End If
        End If
    Next
    
    ' NCJ 24 Jan 03
    If bDoVisitEForm Then
        If Not oEFI.VisitInstance.VisitEFormInstance Is Nothing Then
            ' Loop through each response on the associated Visit eForm (but don't add them to the count)
            ' (responses are automatically loaded by earlier LoadResponses for main eForm)
            For Each oResponse In oEFI.VisitInstance.VisitEFormInstance.Responses
                If oResponse.LockStatus = eLockStatus.lsUnlocked Then
                    If ((bOldStatusMissingRequested) And ((oResponse.Status = eStatus.Missing) Or (oResponse.Status = eStatus.Requested))) _
                        Or ((Not bOldStatusMissingRequested) And (oResponse.Status = eStatus.Unobtainable)) Then
                        'this response is unlocked so toggle this response's status. nb we change derived questions
                        Call oResponse.SetStatusFromSchedule(nToStatus)
                    End If
                End If
            Next
        End If
    End If
    
    If bShowChangeMessage Then
        sMsg = "Are you sure you wish to change " & lChanged & " " & GetStatusText((nOldStatus)) & " response"
        If lChanged > 1 Then
            'if more than one change, make plural
            sMsg = sMsg & "s"
        End If
        sMsg = sMsg & " on '" & oEFI.eForm.Name & "' to " & GetStatusText((nToStatus)) & "?"
        
        'Ask to confirm
        If DialogQuestion(sMsg) = vbYes Then
            bSaveResponses = True
        Else
            bSaveResponses = False
        End If
    Else
        bSaveResponses = True
    End If
            
    ' Save responses
    If bSaveResponses Then
        'save the responses - this will save the subject
        Select Case moSubject.SaveResponses(oEFI, sLockErrMsg)
        Case srrNoLockForSaving
            DialogError sLockErrMsg
        Case srrSubjectReloaded
            ' we'll just try again...
            If moSubject.SaveResponses(oEFI, sLockErrMsg) <> srrSuccess Then
                DialogError "Unable to save changes because another user is editing this subject"
            End If
        Case srrSuccess
            ' OK
        End Select
        If bShowChangeMessage Then
            'ensure schedule is refreshed
            RefreshSchedule
        End If
    End If

    'remove the response from memory
    Call moSubject.RemoveResponses(oEFI, True)
    
Exit Sub

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|modSchedule.ToggleEFIStatus"
    
End Sub

'---------------------------------------------------------------------
Public Function CanOpenEform(oEFI As EFormInstance, _
                            oSubject As StudySubject, oUser As MACROUser, _
                            ByRef sMessage As String) As Boolean
'---------------------------------------------------------------------
' Check if user can open eform they have clicked on
' Returns suitable message in sMessage if result is FALSE
' NCJ 10 Feb 03 - Made Public, and added User & Subject parameters
'---------------------------------------------------------------------
    
    CanOpenEform = False
    
    ' If cannot view data exit
    If Not oUser.CheckPermission(gsFnViewData) Then
        sMessage = "You do not have permission to view subject data"
        Exit Function
    End If
    ' If not a requested status eform then is OK and exit
    If oEFI.Status <> eStatus.Requested Then
        CanOpenEform = True
        Exit Function
    End If
    
    ' Now we have a Requested form
    ' Check permission to change data, subject not readonly
    If Not oUser.CheckPermission(gsFnChangeData) Or oSubject.ReadOnly Then
        sMessage = "You may not enter new data for this subject"
        Exit Function
    End If
    ' Make sure visit is unlocked
    If Not oEFI.VisitInstance.LockStatus = eLockStatus.lsUnlocked Then
        ' NCJ 22 Jan 03 - Change to VisitInstance.LockStatusString when implemented!
        sMessage = "This visit is " & GetLockStatusString(oEFI.VisitInstance.LockStatus) & " and new eForms cannot be opened"
        Exit Function
    End If
    ' NCJ 22 Jan 03 - Since eForm is Requested, must make sure eForm is Unlocked
    If oEFI.LockStatus <> eLockStatus.lsUnlocked Then
        sMessage = "This " & oEFI.LockStatusString & " eForm contains no data and cannot be opened"
        Exit Function
    End If
    
    CanOpenEform = True

End Function

'---------------------------------------------------------------------
Private Function ScheduleUpdatedByOtherUser(ByVal sRedoMsg As String) As Boolean
'---------------------------------------------------------------------
' See if the schedule has been updated by other users
' Returns TRUE if schedule updated OR if there was an update problem,
' or FALSE if nothing's changed and all is OK (so it's safe to continue what we were doing)
' Gives user message if the grid has changed
' sRedoMsg is text for user to tell them to redo what they were doing
' NCJ 22 Aug 06 - Also consider changes to the StudyDef (Bug 2786)
'---------------------------------------------------------------------
Dim sMsg As String
Dim oGrid As ScheduleGrid
Dim sLockErrMsg As String

    On Error GoTo ErrHandler
    
    ScheduleUpdatedByOtherUser = False
    
    ' NCJ 22 Aug 06 - Check study def first
    If CheckStudyDefCurrent(moSubject.StudyId) Then
        ' Study def hasn't changed
        ' Receive all subject updates from other users
        ' Call new StudySubject.Reload routine
        If moSubject.Reload(sLockErrMsg) Then
            ' There's been a reload so refresh the schedule
            Call RefreshSchedule
            sMsg = "This subject has been updated by another user."
            sMsg = sMsg & vbCrLf & sRedoMsg
            DialogInformation sMsg
            ScheduleUpdatedByOtherUser = True
        Else
            ' No reload - was it because of a lock violation?
            ' In this case we don't ask them to try again
            If sLockErrMsg > "" Then
                DialogInformation sLockErrMsg
                ScheduleUpdatedByOtherUser = True
            End If
        End If
    Else
        ' The study definition has changed - reload everything
        ScheduleUpdatedByOtherUser = True
        ' Use top-level subject reload
        Call frmMenu.SubjectOpen(moSubject.StudyId, moSubject.Site, moSubject.PersonId)
        sMsg = "The study has been updated by another user."
        sMsg = sMsg & vbCrLf & sRedoMsg
        DialogInformation sMsg
    End If
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmSchedule.ScheduleUpdatedByOtherUser"

End Function

'---------------------------------------------------------------------
Private Sub ChangeSDVsToDone(oEFI As EFormInstance)
'---------------------------------------------------------------------
' NCJ 12 Aug 04 - Change all Planned question SDVS to Done for this eForm
'---------------------------------------------------------------------
Dim nChanged As Integer
Dim sMsg As String

    On Error GoTo ErrHandler
    
    If DialogQuestion("Are you sure you wish to change all Planned SDVs for questions on this eForm to Done?") = vbYes Then
        ' Make the changes
        nChanged = ChangePlannedQSDVsToDone(moUser.CurrentDBConString, _
                                moUser.UserName, moUser.UserNameFull, _
                                moSubject, oEFI, GetMIMsgSource)
        
        ' Tell them what happened
        Select Case nChanged
        Case 0
            sMsg = "There were no Planned question SDVs on this eForm."
        Case 1
            sMsg = "One SDV was changed to Done."
        Case Else
            sMsg = nChanged & " SDVs were changed to Done."
        End Select
            
        Call DialogInformation(sMsg)
        
        ' Refresh schedule if necessary
        If nChanged > 0 Then
            Call RefreshSchedule
        End If
    End If
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmSchedule.ChangeSDVsToDone"

End Sub

'---------------------------------------------------------------------
Private Sub NewEFormMIMessage(oEFI As EFormInstance, enType As MIMsgType)
'---------------------------------------------------------------------
' Create new MIMessage of given type for this eForm instance
'---------------------------------------------------------------------
Dim bCreateMessage As Boolean
    
    On Error GoTo ErrLabel
    
    bCreateMessage = True
    
    ' Check to see if we already have an SDV for this eForm
    If enType = MIMsgType.mimtSDVMark And oEFI.SDVStatus <> ssNone Then
        If SDVExists(MIMsgScope.mimscEForm, _
                        moSubject.StudyCode, moSubject.Site, moSubject.PersonId, _
                        oEFI.VisitInstance.VisitId, oEFI.VisitInstance.CycleNo, _
                        oEFI.EFormTaskId) Then
            ' MLM 01/07/05: Show the existing SDV instead of giving an error
            ShowMIMessage MIMsgType.mimtSDVMark, oEFI.eForm.Study.StudyId, moSubject.Site, moSubject.PersonId, oEFI.VisitInstance.VisitId, oEFI.VisitInstance.CycleNo, oEFI.eForm.EFormId, oEFI.CycleNo
'            DialogInformation "An SDV Mark already exists for this eForm"
            bCreateMessage = False
        End If
    End If
    
    If bCreateMessage Then
        With moSubject
            If CreateNewMIMessage(enType, MIMsgScope.mimscEForm, _
                    IMedNow, moSubject.TimeZone.TimezoneOffset, _
                    .Site, .StudyId, .PersonId, Nothing, _
                    oEFI.VisitInstance.VisitId, oEFI.VisitInstance.CycleNo, _
                    oEFI.EFormTaskId, _
                    0, 0, "", 0, _
                    oEFI.eForm.EFormId, oEFI.CycleNo, 0, "", _
                    moSubject) Then
                    'datausername is blank for eforms
                Call UpdateMIMsgStatus(gsADOConnectString, enType, _
                        .StudyDef.Name, .StudyDef.StudyId, .Site, .PersonId, _
                        oEFI.VisitInstance.VisitId, oEFI.VisitInstance.CycleNo, oEFI.EFormTaskId, _
                        0, 0, moSubject)
                        
                'ensure schedule is refreshed
                RefreshSchedule
            End If
        End With
    End If
    
Exit Sub
ErrLabel:
    If MACROErrorHandler("modSchedule", Err.Number, Err.Description, "NewEFormMIMessage", Err.Source) = Retry Then
        Resume
    End If

End Sub

'---------------------------------------------------------------------
Private Sub NewVisitMIMessage(oVI As VisitInstance, enType As MIMsgType)
'---------------------------------------------------------------------
' Create new MIMessage of given type for this Visit instance
'---------------------------------------------------------------------
Dim bCreateMessage As Boolean
    
    On Error GoTo ErrLabel
    
    bCreateMessage = True
    
    ' Check to see if we already have an SDV for this visit
    If enType = MIMsgType.mimtSDVMark And oVI.SDVStatus <> ssNone Then
        If SDVExists(MIMsgScope.mimscVisit, _
                        moSubject.StudyCode, moSubject.Site, moSubject.PersonId, _
                        oVI.VisitId, oVI.CycleNo) Then
            ' MLM 01/07/05: Show the existing SDV instead of giving an error
            ShowMIMessage MIMsgType.mimtSDVMark, moSubject.StudyId, moSubject.Site, moSubject.PersonId, oVI.VisitId, oVI.CycleNo
            'DialogInformation "An SDV Mark already exists for this visit"
            bCreateMessage = False
        End If
    End If
    
    If bCreateMessage Then
        With moSubject
            If CreateNewMIMessage(enType, MIMsgScope.mimscVisit, _
                    IMedNow, moSubject.TimeZone.TimezoneOffset, _
                    .Site, .StudyId, .PersonId, Nothing, _
                    oVI.VisitId, oVI.CycleNo, _
                    0, 0, 0, "", 0, _
                    0, 0, 0, "", _
                    moSubject) Then
                    'datausername is blank for eforms
                Call UpdateMIMsgStatus(gsADOConnectString, enType, _
                        .StudyDef.Name, .StudyDef.StudyId, .Site, .PersonId, _
                        oVI.VisitId, oVI.CycleNo, _
                        0, 0, 0, moSubject)
            
                'ensure schedule is refreshed
                RefreshSchedule
            End If
        End With
    End If
    
Exit Sub
ErrLabel:
    If MACROErrorHandler("modSchedule", Err.Number, Err.Description, "NewVisitMIMessage", Err.Source) = Retry Then
        Resume
    End If

End Sub

'---------------------------------------------------------------------
Private Sub NewSubjectMIMessage(oSubject As StudySubject, enType As MIMsgType)
'---------------------------------------------------------------------
' Create new MIMessage of given type for this subject
'---------------------------------------------------------------------
Dim bCreateMessage As Boolean

    On Error GoTo ErrLabel
    
    bCreateMessage = True
    
    ' Check to see if we already have an SDV for this subject
    If enType = MIMsgType.mimtSDVMark And oSubject.SDVStatus <> ssNone Then
        If SDVExists(MIMsgScope.mimscSubject, _
                        oSubject.StudyCode, oSubject.Site, oSubject.PersonId) Then
            ' MLM 01/07/05: Show the existing SDV instead of giving an error
            ShowMIMessage MIMsgType.mimtSDVMark, oSubject.StudyId, oSubject.Site, oSubject.PersonId
            'DialogInformation "An SDV Mark already exists for this subject"
            bCreateMessage = False
        End If
    End If
    
    If bCreateMessage Then
        With moSubject
            If CreateNewMIMessage(enType, MIMsgScope.mimscSubject, _
                    IMedNow, moSubject.TimeZone.TimezoneOffset, _
                    .Site, .StudyId, .PersonId, Nothing, _
                    0, 0, 0, 0, 0, "", 0, _
                    0, 0, 0, "", _
                    moSubject) Then
                    'datausername is blank for eforms
                Call UpdateMIMsgStatus(gsADOConnectString, enType, _
                        .StudyDef.Name, .StudyDef.StudyId, .Site, .PersonId, _
                        0, 0, 0, 0, 0, moSubject)
            
                'ensure schedule is refreshed
                RefreshSchedule
            End If
        End With
    End If
    
Exit Sub
ErrLabel:
    If MACROErrorHandler("modSchedule", Err.Number, Err.Description, "NewSubjectMIMessage", Err.Source) = Retry Then
        Resume
    End If

End Sub

'---------------------------------------------------------------------
Private Sub ShowMIMessage(enMIMsgType As MIMsgType, lStudyId As Long, sSite As String, lPersonId As Long, _
    Optional lVisitId As Long = -1, Optional lVisitCycle As Long = -1, _
    Optional lEFormId As Long = -1, Optional lEFormCycle As Long = -1)
'---------------------------------------------------------------------
' MLM 29/06/05: Added: Launch the MIMessage browser for the selected item.
'---------------------------------------------------------------------
Dim sStatus As String
Dim vData As Variant
Dim sScope As String
Dim ofrmMIMModal As Form

    On Error GoTo ErrLabel
    
    Select Case enMIMsgType
    Case MIMsgType.mimtDiscrepancy: sStatus = "111" 'three statuses of discrepancy
    Case MIMsgType.mimtNote: sStatus = "11" 'two statuses of note
    Case MIMsgType.mimtSDVMark: sStatus = "1111" 'four statuses of SDV
    End Select
    
    'specify that we only want MIMessages on the selected object, not its parents or children
    If lVisitId = -1 Then
        sScope = "1000" 'subject-level
    ElseIf lEFormId = -1 Then
        sScope = "0100" 'visit-level
    Else
        sScope = "0010" 'eform-level
    End If
    
    vData = GetWinIO.GetMIMessageList(enMIMsgType, goUser, lStudyId, sSite, _
                              lVisitId, lVisitCycle, lEFormId, lEFormCycle, "", "", "", "", lPersonId, _
                                sStatus, "", "", sScope)
                                                       
    If Not IsNull(vData) Then
       Set ofrmMIMModal = frmViewDiscrepanciesModal
        Call ofrmMIMModal.DisplayModal(enMIMsgType, SubjectMIMEssage, vData)
        Set ofrmMIMModal = Nothing
        
        RefreshSchedule
    Else
        MsgBox "No matching records"
    End If
    Exit Sub
    
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modSchedule.ShowMIMessage"
End Sub

'---------------------------------------------------------------------
Private Sub ChangeLockUnlockFreeze(ByVal enAction As LFAction, enScope As LFScope, _
                lStudyId As Long, sStudyName As String, sSite As String, lSubjectId As Long, _
                Optional lVisitId As Long = 0, Optional nVisitCycleNumber As Integer = 0, _
                Optional lCRFPageTaskId As Long = 0, _
                Optional lCRFPageID As Long = 0, Optional nCRFPageCycleNumber As Integer = 0)
'---------------------------------------------------------------------
'   SDM 07/12/99
' CHange the lock status of an object
' TA 5/10/2201: Subject is now locked during the update
' ATO 20/08/2002 Added RepeatNumber
' NCJ 23 Dec 02 - Now takes LockFreeze Action rather than LockStatus
' NCJ 3 Jan 03 - Use QuestionId too
' NCJ 9 Jan03 - Do a Refresh for Unfreeze (because we don't know what the final statuses are going to be)
' DPH 14/01/2003 - Copied and adapted from frmDataItemResponse version
'---------------------------------------------------------------------
Dim sTimestamp As String 'ATO 21/08/2002

Dim sMsg As String
Dim sToken As String
Dim lMVDataIndex As Long

' NCJ 23 Dec 02 - New Lock/Freeze objects
Dim oLFObj As LFObject
Dim oFlocker As LockFreeze
Dim oLF As LockFreeze
Dim nSource As Integer
Dim enStatus As LockStatus

    On Error GoTo ErrLabel
    
    ' check if can Lock/Freeze
    Set oLF = New LockFreeze
    If Not gblnRemoteSite And Not oLF.CanLockFreezeOnServer(MacroADODBConnection, _
                                    sStudyName, sSite, lSubjectId) Then
        DialogInformation "Lock/freeze operations may not be carried out on this subject because there is unimported site data"
        Exit Sub
    End If
    Set oLF = Nothing
    
    'TA 27/9/01: this code checks to see if the subject is locked
    'nb The subject could be locked by by the current user having the schedule open
    sToken = LockSubject(goUser.UserName, lStudyId, sSite, lSubjectId)
    If sToken = "" Then
        ' this subject is currently locked (Message already given to User in LockSubject)
        Exit Sub
    End If
    
    ' NCJ 23 Dec 02 - You CAN Unfreeze in MACRO 3.0!
    HourglassOn
    
    Set oLFObj = New LFObject
    Select Case enScope
    Case LFScope.lfscSubject
        ' Apply to the Trial Subject
        Call oLFObj.Init(LFScope.lfscSubject, lStudyId, sSite, lSubjectId)
    Case LFScope.lfscVisit
        ' Apply to the Visit
        Call oLFObj.Init(LFScope.lfscVisit, lStudyId, sSite, lSubjectId, _
                        lVisitId, nVisitCycleNumber)
    Case LFScope.lfscEForm
        ' Apply to the eForm
        Call oLFObj.Init(LFScope.lfscEForm, lStudyId, sSite, lSubjectId, _
                        lVisitId, nVisitCycleNumber, _
                        lCRFPageID, nCRFPageCycleNumber)
    End Select
    
    ' NCJ 23 Dec 02 - Get the LockFreeze object to do all the work
    Set oFlocker = New LockFreeze
    If gblnRemoteSite Then
        nSource = TypeOfInstallation.RemoteSite
    Else
        nSource = TypeOfInstallation.Server
    End If
    Call oFlocker.DoLockFreeze(MacroADODBConnection, oLFObj, enAction, nSource, _
                                goUser.UserName, goUser.UserNameFull)
    
    ' unlock the subject (this is the database lock!)
    UnlockSubject lStudyId, sSite, lSubjectId, sToken
    
    ' Refresh subject in memory - close then reopen
    CloseSubject moSubject.StudyDef, False, False, True
    
    frmMenu.SubjectOpen lStudyId, sSite, lSubjectId
        
    HourglassOff
    
    Set oLFObj = Nothing
    Set oFlocker = Nothing
        
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|modSchedule.ChangeLockUnlockFreeze"

End Sub

'--------------------------------------------------------------
Private Sub ToggleVisitStatus(oVI As VisitInstance, nToStatus As eStatus, _
                            Optional bShowChangeMessage As Boolean = True)
'--------------------------------------------------------------
' Make a Visit unobtainable by changing all its 'Missing'/'Requested' responses to 'Unobtainable'
' or vice versa
' no permissions etc are checked in here as they were checked to enable the menu items
' NCJ 23/24 Jan 03 - Must make sure we deal with the VisitEForm correctly!
'--------------------------------------------------------------
Dim oEFI As EFormInstance
Dim lChanged As Long
Dim nOldStatus As eStatus
Dim sStatus As String
Dim sMsg As String
Dim sLockErrMsg As String
Dim bSaveResponses As Boolean
Dim bOldStatusMissingRequested As Boolean
Dim lVEFITaskId As Long
Dim i As Long

    On Error GoTo ErrLabel
    
    ' NCJ 26 Sept 02 - Check for subject data updates first
    ' and don't continue if there are any
    If ScheduleUpdatedByOtherUser("Please try again") Then
        HourglassOff
        Exit Sub
    End If
    
    If (nToStatus = eStatus.Missing) Then
        nOldStatus = eStatus.Unobtainable
        bOldStatusMissingRequested = False
    Else
        nOldStatus = eStatus.Missing
        bOldStatusMissingRequested = True
    End If
    
    ' NCJ 24 Jan 03 - Get the Visit eForm, if any
    lVEFITaskId = 0
    If Not oVI.VisitEFormInstance Is Nothing Then
        lVEFITaskId = oVI.VisitEFormInstance.EFormTaskId
    End If
    
    If bShowChangeMessage Then
        lChanged = 0
        'loop through eforms checking how many will be changed (don't count Visit eForm)
        For Each oEFI In oVI.eFormInstances
            If oEFI.LockStatus = eLockStatus.lsUnlocked And (oEFI.EFormTaskId <> lVEFITaskId) Then
                If ((bOldStatusMissingRequested) And ((oEFI.Status = eStatus.Missing) Or (oEFI.Status = eStatus.Requested))) _
                    Or ((Not bOldStatusMissingRequested) And (oEFI.Status = eStatus.Unobtainable)) Then
                    'increase the count of eforms that will be changed
                    lChanged = lChanged + 1
                End If
            End If
        Next
        
        ' Are there any to change?
        If lChanged > 0 Then
            sMsg = "Are you sure you wish to set " & lChanged & " eForm"
            If lChanged > 1 Then
                'if more than one change, make plural
                sMsg = sMsg & "s"
            End If
            sMsg = sMsg & " in Visit '" & oVI.Name & "' to " & GetStatusText((nToStatus)) & "?"
            
            'Ask to confirm
            If DialogQuestion(sMsg) = vbYes Then
                bSaveResponses = True
            Else
                bSaveResponses = False
            End If
        Else
            ' No eForms to change
            DialogInformation "There are no eForms within this visit that can be changed to " & GetStatusText((nToStatus))
            bSaveResponses = False
        End If
    Else
        bSaveResponses = True
    End If
            
    ' Save responses
    If bSaveResponses Then
    
        If bShowChangeMessage Then
            Call frmHourglass.Display("Changes being saved...", True)
        End If
        
        ' loop through eForms calling ToggleEFIStatus
        ' NCJ 24 Jan 03 - Must deal with VisitEForm with first eForm processed
        For i = 1 To oVI.eFormInstances.Count
            Set oEFI = oVI.eFormInstances(i)
            If oEFI.LockStatus = eLockStatus.lsUnlocked And (oEFI.EFormTaskId <> lVEFITaskId) Then
                If ((bOldStatusMissingRequested) And ((oEFI.Status = eStatus.Missing) Or (oEFI.Status = eStatus.Requested))) _
                    Or ((Not bOldStatusMissingRequested) And (oEFI.Status = eStatus.Unobtainable)) Then
                        ' Also process VisitEForm if i=1
                        Call ToggleEFIStatus(oEFI, nToStatus, False, (i = 1))
                End If
            End If
        Next i
        
        If bShowChangeMessage Then
            'ensure schedule is refreshed
            RefreshSchedule
        End If
    End If
    
    If bShowChangeMessage Then
        UnloadfrmHourglass
    End If

Exit Sub

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|modSchedule.ToggleVisitStatus"
End Sub

'--------------------------------------------------------------
Private Sub ToggleSubjectStatus(oSubject As StudySubject, nToStatus As eStatus)
'--------------------------------------------------------------
' Make a Subject unobtainable by changing all its 'Missing'/'Requested' responses to 'Unobtainable'
' or vice versa
' no permissions etc are checked in here as they were checked to enable the menu items
'--------------------------------------------------------------
Dim oVI As VisitInstance
Dim lChanged As Long
Dim nOldStatus As eStatus
Dim sStatus As String
Dim sMsg As String
Dim sLockErrMsg As String
Dim sLabel As String
Dim bOldStatusMissingRequested As Boolean

    On Error GoTo ErrLabel
    
    
    ' NCJ 26 Sept 02 - Check for subject data updates first
    ' and don't continue if there are any
    If ScheduleUpdatedByOtherUser("Please try again") Then
        Exit Sub
    End If
    
    If (nToStatus = eStatus.Missing) Then
        nOldStatus = eStatus.Unobtainable
        bOldStatusMissingRequested = False
    Else
        nOldStatus = eStatus.Missing
        bOldStatusMissingRequested = True
    End If
    
    lChanged = 0
    'loop through visits checking how many will be changed
    For Each oVI In oSubject.VisitInstances
        If oVI.LockStatus = eLockStatus.lsUnlocked Then
            If ((bOldStatusMissingRequested) And ((oVI.Status = eStatus.Missing) Or (oVI.Status = eStatus.Requested))) _
                Or ((Not bOldStatusMissingRequested) And (oVI.Status = eStatus.Unobtainable)) Then
                    'increase the count of eforms that will be changed
                    lChanged = lChanged + 1
            End If
        End If
    Next

    ' Check to see if there are any changes to make
    If lChanged > 0 Then
        sMsg = "Are you sure you wish to change " & lChanged & " " & GetStatusText((nOldStatus)) & " Visit"
        If lChanged > 1 Then
            'if more than one change, make plural
            sMsg = sMsg & "s"
        End If
        sLabel = oSubject.label
        If sLabel = "" Then
            sLabel = oSubject.PersonId
        End If
        sLabel = oSubject.StudyCode & "/" & oSubject.Site & "/" & sLabel
        sMsg = sMsg & " in Subject " & sLabel & " to " & GetStatusText((nToStatus)) & "?"
            
        'Ask to confirm
        If DialogQuestion(sMsg) = vbYes Then
            Call frmHourglass.Display("Changes being saved...", True)
    
            ' loop through Visits calling ToggleVisitStatus
            For Each oVI In oSubject.VisitInstances
                If oVI.LockStatus = eLockStatus.lsUnlocked Then
                    If ((bOldStatusMissingRequested) And ((oVI.Status = eStatus.Missing) Or (oVI.Status = eStatus.Requested))) _
                        Or ((Not bOldStatusMissingRequested) And (oVI.Status = eStatus.Unobtainable)) Then
                            Call ToggleVisitStatus(oVI, nToStatus, False)
                    End If
                End If
            Next
            
            'ensure schedule is refreshed
            RefreshSchedule
        End If
    Else
        ' Nothing to change
        Call DialogInformation("There are no visits for this subject than can be changed to " & GetStatusText((nToStatus)))
    End If
    
    UnloadfrmHourglass

Exit Sub

ErrLabel:

    Err.Raise Err.Number, , Err.Description & "|modSchedule.ToggleSubjectStatus"
End Sub

'--------------------------------------------------------------
Private Function CanMakeVisitUnobtainable(oVI As VisitInstance) As Boolean
'--------------------------------------------------------------
' Returns TRUE if current user can unobtainablise the given Visit instance
'--------------------------------------------------------------
' REVISIONS
'--------------------------------------------------------------
    CanMakeVisitUnobtainable = False
    
    ' User must be able to change data
    If Not goUser.CheckPermission(gsFnChangeData) Then Exit Function
    
    ' Subject must not be Read-Only
    If moSubject.ReadOnly Then Exit Function
    
    ' Form must not be locked or frozen
    If oVI.LockStatus <> eLockStatus.lsUnlocked Then Exit Function
    
    ' Form must have status of missing OR requested
    If ((oVI.Status <> eStatus.Missing) And (oVI.Status <> eStatus.Requested)) Then Exit Function
    
    ' If we get here, all is OK
    CanMakeVisitUnobtainable = True

End Function

'--------------------------------------------------------------
Private Function CanMakeSubjectUnobtainable() As Boolean
'--------------------------------------------------------------
' Returns TRUE if current user can unobtainablise the given Subject
'--------------------------------------------------------------
' REVISIONS
'--------------------------------------------------------------
    CanMakeSubjectUnobtainable = False
    
    ' User must be able to change data
    If Not goUser.CheckPermission(gsFnChangeData) Then Exit Function
    
    ' Subject must not be Read-Only
    If moSubject.ReadOnly Then Exit Function
    
    ' Form must not be locked or frozen
    If moSubject.LockStatus <> eLockStatus.lsUnlocked Then Exit Function
    
    ' Form must have status of missing OR requested
    If ((moSubject.Status <> eStatus.Missing) And (moSubject.Status <> eStatus.Requested)) Then Exit Function
    
    ' If we get here, all is OK
    CanMakeSubjectUnobtainable = True

End Function

'--------------------------------------------------------------
Private Function CanMakeVisitMissing(oVI As VisitInstance) As Boolean
'--------------------------------------------------------------
' Returns TRUE if current user can make missing the given Visit instance
'--------------------------------------------------------------

    CanMakeVisitMissing = False
    
    ' User must be able to change data
    If Not goUser.CheckPermission(gsFnChangeData) Then Exit Function
    
    ' Subject must not be Read-Only
    If moSubject.ReadOnly Then Exit Function
    
    ' Form must not be locked or frozen
    If oVI.LockStatus <> eLockStatus.lsUnlocked Then Exit Function
    
    ' Form must have status of unobtanable
    If oVI.Status <> eStatus.Unobtainable Then Exit Function

    ' If we get here, all is OK
    CanMakeVisitMissing = True

End Function

'--------------------------------------------------------------
Private Function CanMakeSubjectMissing() As Boolean
'--------------------------------------------------------------
' Returns TRUE if current user can make missing the given eForm instance
'--------------------------------------------------------------

    CanMakeSubjectMissing = False
    
    ' User must be able to change data
    If Not goUser.CheckPermission(gsFnChangeData) Then Exit Function
    
    ' Subject must not be Read-Only
    If moSubject.ReadOnly Then Exit Function
    
    ' Form must not be locked or frozen
    If moSubject.LockStatus <> eLockStatus.lsUnlocked Then Exit Function
    
    ' Form must have status of unobtanable
    If moSubject.Status <> eStatus.Unobtainable Then Exit Function

    ' If we get here, all is OK
    CanMakeSubjectMissing = True

End Function


