Attribute VB_Name = "modSCM"

'----------------------------------------------------------------------------
'   File:       modSCM.bas
'   Copyright:  InferMed Ltd. 2002-2004. All Rights Reserved
'   Author:     Nicky Johns, July 2002
'   Purpose:    Handling for CachedSubject objects for the Cache Manager
'-----------------------------------------------------------------------------
' Revisions:
'   NCJ 18-22 July - Initial development
'   NCJ 23 July 02 - Read prolog memory settings from registry in InitSCM
'               Some tidying up (following code review by TA)
'   MACRO 3.0
'   NCJ 17 Sept 02 - Updated for new MACRO 3.0 locking model
'   NCJ 24 Jan 03 - Added sCountry parameter to NewSubject
'   NCJ 29 Jan 03 - Set up msPrologSwitches the first time it's needed, using clsArezzoMemory
'   NCJ 29 May 03 - Added debugging to record any suspect BUSY cache entries
'   ic 18/08/2003 added LoadEformB() and SaveEformB() to improve performance on certain machines
'   ic 15/09/2003 added CheckArezzoEventsB() and SaveArezzoEventsB() to improve performance on certain machines
'   NCJ 15 Jan 04 - Changed separator from comma to pipe in Cache Report
'   ic 16/03/2004 remove conditional ORAMA compilation
'   ic 20/12/2004 bug 2395 - pass the sSerialisedUser byref so that lastused eform gets set and passed back
'   ic 18/05/2005 issue 2560, compare database connection when locating cached subjects
'----------------------------------------------------------------------------

Option Explicit

' User defined types for subject states
Public Enum eBusyStatus
    SubjectBusy = 0
    SubjectNotBusy = 1
End Enum

' The collection of CachedSubject objects
Private mcolSCMSubjects As Collection

' Used for generating unique but otherwise unimportant keys
Private mnKeyValue As Integer

' The memory settings to use for Prolog
Private msPrologSwitches As String

'----------------------------------------------------------------------------
Public Sub InitSCM()
'----------------------------------------------------------------------------
' Set things up before we start
' if we haven't already done it
' revisions
' ic 28/01/2003 changed to get max arezzo from settings file
'----------------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    If mcolSCMSubjects Is Nothing Then
        Set mcolSCMSubjects = New Collection
        
        ' Start with key no. 1
        mnKeyValue = 1
        
        ' Read no. of cache entries from registry
        'gnMaxArezzoAllowed = Val(GetRegVal("MaxSubjectsAllowed"))
        ' If there was no registry setting, default to 1
        'If gnMaxArezzoAllowed < 1 Then
        '    gnMaxArezzoAllowed = 1
        'End If
        InitialiseSettingsFile True
        gnMaxArezzoAllowed = CInt(GetMACROSetting(MACRO_SETTING_MAXAREZZO, "1"))
        
        ' Read Prolog memory settings from registry
        'msPrologSwitches = GetPrologSwitches
        
        ' NCJ 29 Jan 03 - Prolog switches will be read in the first time they're needed
        msPrologSwitches = ""
    
    End If

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "| modSCM.InitSCM"

End Sub

'----------------------------------------------------------------------------
Public Function CreateNewSubjectB(ByVal sUserName As String, _
                                 ByVal sDatabaseCnn As String, _
                                 ByVal lStudyId As Long, _
                                 ByVal sSite As String, _
                                 ByVal sCountry As String, _
                                ByVal sUserNameFull As String, _
                                ByVal sUserRole As String) As StudySubject
'----------------------------------------------------------------------------
' Create a new subject and return the StudySubject object
' May be Nothing if we couldn't get a cache object
' NCJ 24 Jan 03 - Added sCountry parameter
' NCJ 6 May 03 - Added new params. sUserNameFull and sUserRole
'----------------------------------------------------------------------------
Dim oCSubject As CachedSubject

    ' NCJ 29 Jan 03 - Ensure the Prolog switches are set up
    Call GetPrologSwitches(lStudyId, sDatabaseCnn)
    
    ' Get ourselves a usable Cache Object
    Set oCSubject = GetUsableCacheObject
    If Not oCSubject Is Nothing Then
        ' We got one
        Set CreateNewSubjectB = oCSubject.NewSubject(lStudyId, sSite, sUserName, sCountry, sDatabaseCnn, _
                            sUserNameFull, sUserRole)
    End If
    
    ' NCJ 29 May 03 - Report on suspect BUSY entries
    Call ReportBusyCacheEntries

End Function

'----------------------------------------------------------------------------
Public Function ReleaseSubjectB(ByVal sDatabaseCnn As String, ByVal lStudyId As Long, _
                            ByVal sSite As String, _
                            ByVal lSubjectId As Long, _
                            ByVal sToken As String) As Boolean
'----------------------------------------------------------------------------
' Release a specific subject, i.e. mark it as "Not In Use"
' Returns TRUE if it was OK,
' or FALSE if we couldn't find the subject or if it wasn't Busy
' ic 18/05/2005 issue 2560, pass database connection string
'----------------------------------------------------------------------------
Dim oCSubject As CachedSubject
Dim bReleased As Boolean

    On Error GoTo ErrLabel
    
    bReleased = False
    
    ' Look for Cache entry with this token, with status Busy
    Set oCSubject = GetCachedSubject(sDatabaseCnn, lStudyId, sSite, lSubjectId, sToken, SubjectBusy)
    ' We ought to have one, but do nothing if we didn't find it
    If Not oCSubject Is Nothing Then
        If oCSubject.BusyStatus = SubjectBusy Then
            Call oCSubject.ReleaseSubject
            bReleased = True
        End If
        Set oCSubject = Nothing
    End If
    
    ReleaseSubjectB = bReleased

    ' NCJ 29 May 03 - Report on suspect BUSY entries
    Call ReportBusyCacheEntries

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "| modSCM.ReleaseSubjectB( " _
                    & lStudyId & ", " & sSite & ", " & lSubjectId & " )"

End Function

'----------------------------------------------------------------------------
Public Function LoadSubjectB(ByVal sDatabaseCnn As String, _
                        ByVal lStudyId As Long, _
                        ByVal sSite As String, _
                        ByVal lSubjectId As Long, _
                        ByVal sUserName As String, _
                        ByVal nUpdateMode As Integer, _
                        ByVal sUserNameFull As String, _
                        ByVal sUserRole As String) As StudySubject
'----------------------------------------------------------------------------
' Load the specified subject, and return StudySubject object
' Returns Nothing if we couldn't do it
' NCJ 6 May 03 - Added new params. sUserNameFull and sUserRole
' ic 18/05/2005 issue 2560, pass database connection string
'----------------------------------------------------------------------------
Dim oCSubject As CachedSubject

    On Error GoTo ErrLabel
    
    ' NCJ 29 Jan 03 - Ensure the Prolog switches are set up
    Call GetPrologSwitches(lStudyId, sDatabaseCnn)
    
    ' See if we have it loaded already (but ignoring its token)
    Set oCSubject = GetCachedSubject(sDatabaseCnn, lStudyId, sSite, lSubjectId, "", SubjectNotBusy)
    
    If Not oCSubject Is Nothing Then
        If oCSubject.BusyStatus = SubjectBusy Then
            ' Shouldn't happen - return Nothing
            Set LoadSubjectB = Nothing
        Else
            ' We have this subject already - get it to reload if necessary
            Set LoadSubjectB = oCSubject.Reload(sUserName, nUpdateMode, sUserNameFull, sUserRole)
        End If
    Else
        ' We don't have this subject already
        ' so get a cache object we can use
        Set oCSubject = GetUsableCacheObject
        
        If Not oCSubject Is Nothing Then
            ' Load the subject into it
            Set LoadSubjectB = oCSubject.Load(lStudyId, sSite, lSubjectId, sUserName, nUpdateMode, sDatabaseCnn, _
                            sUserNameFull, sUserRole)
        Else
            ' No cache objects available
            Set LoadSubjectB = Nothing
        End If
    End If
    
    ' Tidy up
    Set oCSubject = Nothing
    
    ' NCJ 29 May 03 - Report on suspect BUSY entries
    Call ReportBusyCacheEntries
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "| modSCM.LoadSubjectB( " _
                    & lStudyId & ", " & sSite & ", " & lSubjectId & " )"

End Function

'----------------------------------------------------------------------------
Private Function GetCachedSubject(sDatabaseCnn As String, lStudyId As Long, sSite As String, _
                            lSubjectId As Long, _
                            sToken As String, _
                            enBusyStatus As eBusyStatus) As CachedSubject
'----------------------------------------------------------------------------
' Find a CachedSubject which matches this specified MACRO subject
' and return it if found
' If sToken is given, must match on Token too
' If bAvailable = TRUE, search for a not busy cache entry
' Returns Nothing if no matching subject
' ic 18/05/2005 issue 2560, pass database connection string
'----------------------------------------------------------------------------
Dim i As Integer
Dim oCSubject As CachedSubject

    On Error GoTo ErrLabel
    
    If mcolSCMSubjects Is Nothing Then
        ' There aren't any!
        Call InitSCM
        Set GetCachedSubject = Nothing
        Exit Function
    End If
    
    ' Search for specified subject
    For i = 1 To mcolSCMSubjects.Count
        Set oCSubject = mcolSCMSubjects(i)
        If CachedSubjectMatches(oCSubject, sDatabaseCnn, lStudyId, sSite, lSubjectId, sToken, enBusyStatus) Then
            ' We've found one that matches
            Set GetCachedSubject = oCSubject
            Exit For
        End If
    Next i
    
    Set oCSubject = Nothing

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "| modSCM.GetCachedSubject( " _
                    & lStudyId & ", " & sSite & ", " & lSubjectId & " )"

End Function

'----------------------------------------------------------------------------
Private Function CachedSubjectMatches(oCSubject As CachedSubject, sDatabaseCnn As String, _
                            lStudyId As Long, _
                            sSite As String, _
                            lSubjectId As Long, _
                            sToken As String, _
                            enBusyStatus As eBusyStatus) As Boolean
'----------------------------------------------------------------------------
' NCJ 17 Sept 02
' Returns TRUE if the CachedSubject matches on Study, Site, Subject
' and Token (if non-empty) and BusyStatus
' or FALSE otherwise
' ic 18/05/2005 issue 2560, compare database connection string too
'----------------------------------------------------------------------------

    CachedSubjectMatches = False
    
    If oCSubject.DatabaseCnn <> sDatabaseCnn Then Exit Function
    If oCSubject.StudyId <> lStudyId Then Exit Function
    If oCSubject.Site <> sSite Then Exit Function
    If oCSubject.SubjectId <> lSubjectId Then Exit Function
    
    If oCSubject.BusyStatus <> enBusyStatus Then Exit Function
    
    ' See if we want a match on token
    If sToken > "" Then
        If oCSubject.ArezzoToken <> sToken Then Exit Function
    End If
        
    ' If we get here there are no mismatches
    CachedSubjectMatches = True

End Function

'----------------------------------------------------------------------------
Private Function GetUsableCacheObject() As CachedSubject
'----------------------------------------------------------------------------
' Find a CachedSubject which can be used for a new subject
' and return it if found (we don't do any matching of subjects here)
' Returns Nothing if all are busy
'----------------------------------------------------------------------------

    If mcolSCMSubjects.Count < gnMaxArezzoAllowed Then
        ' We haven't reached the max yet so just create a new one
        Set GetUsableCacheObject = CreateNewCachedSubject
    Else
        ' Go and get the oldest one that's not busy
        Set GetUsableCacheObject = GetOldestUnbusyCacheObject
    End If

End Function

'----------------------------------------------------------------------------
Private Function GetOldestUnbusyCacheObject() As CachedSubject
'----------------------------------------------------------------------------
' Get the least recently used Cache object that's not busy
' Returns Nothing if they're all busy
'----------------------------------------------------------------------------
Dim i As Integer
Dim oCSubject As CachedSubject
Dim oCFoundSubject As CachedSubject
Dim dblTime As Double

    On Error GoTo ErrLabel
    
    ' Initialise time stamp to current time
    dblTime = CDbl(Now)
    ' We haven't found one yet
    Set oCFoundSubject = Nothing
    
    ' We loop through the cache objects and find
    ' the "least recently used" one that's not busy
    For i = 1 To mcolSCMSubjects.Count
        Set oCSubject = mcolSCMSubjects(i)
        ' Ignore Busy ones
        If oCSubject.BusyStatus <> SubjectBusy Then
            ' It's not busy
            If oCSubject.TimeStamp <= dblTime Then
                ' Its timestamp is earlier so go for this one
                Set oCFoundSubject = oCSubject
                dblTime = oCSubject.TimeStamp
            End If
        End If
    Next i

    ' oCFoundSubject might still be Nothing
    Set GetOldestUnbusyCacheObject = oCFoundSubject

    ' Tidy up our objects
    Set oCSubject = Nothing
    Set oCFoundSubject = Nothing

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "| modSCM.GetOldestUnbusyCacheObject"

End Function

'----------------------------------------------------------------------------
Public Function ReduceCacheObjects(ByVal nTargetNo As Integer) As Integer
'----------------------------------------------------------------------------
' Reduce number of cache objects to nTargetNo if we can
' Returns how many are actually left when we've finished
' (We can't remove Busy ones)
'----------------------------------------------------------------------------
Dim nRemoved As Integer
Dim nToRemove As Integer
Dim oCSubject As CachedSubject

    ' We want to remove enough to reach nTargetNo
    nToRemove = gnMaxArezzoAllowed - nTargetNo
    
    ' We may not have created more than the target
    If mcolSCMSubjects.Count <= nTargetNo Then
        ReduceCacheObjects = nTargetNo
        Exit Function
    End If
    
    nRemoved = 0
    
    Do While nRemoved < nToRemove
        ' Get the oldest not busy one
        Set oCSubject = GetOldestUnbusyCacheObject
        If oCSubject Is Nothing Then
            ' We can't do any more
            Exit Do
        Else
            ' We shut it down and remove it from the collection
            Call mcolSCMSubjects.Remove(oCSubject.MyKey)
            Call oCSubject.Terminate
            Set oCSubject = Nothing
            nRemoved = nRemoved + 1
        End If
    Loop
    
    ReduceCacheObjects = nTargetNo - nRemoved

End Function

'----------------------------------------------------------------------------
Private Function CreateNewCachedSubject() As CachedSubject
'----------------------------------------------------------------------------
' Create a new object and add it to the collection
' NB We generate unique sequential keys
' NB Assume msPrologSwitches already set up
'----------------------------------------------------------------------------
Dim oCSubject As CachedSubject
Dim sKey As String

    Set oCSubject = New CachedSubject
    sKey = NewKey
    Call oCSubject.Init(sKey, App.Path & "\Temp\", msPrologSwitches)
    ' Add it to our collection
    mcolSCMSubjects.Add oCSubject, sKey
    
    Set CreateNewCachedSubject = oCSubject
    
    Set oCSubject = Nothing

End Function

'----------------------------------------------------------------------------
Public Sub CloseSCM()
'----------------------------------------------------------------------------
' Tidy things up before we go home
' NB This will close down all objects and terminate all Prolog instances
'----------------------------------------------------------------------------
Dim i As Integer
Dim oCSubject As CachedSubject

    If Not mcolSCMSubjects Is Nothing Then
        ' Close down each CachedSubject object
        For i = 1 To mcolSCMSubjects.Count
            Set oCSubject = mcolSCMSubjects(i)
            ' Terminate includes closing down its Prolog
            Call oCSubject.Terminate
        Next i
        Set oCSubject = Nothing
        Set mcolSCMSubjects = Nothing
    End If
    
End Sub

'----------------------------------------------------------------------------
Public Function GetSubjectStatusString(ByVal nBusyStatus As eBusyStatus) As String
'----------------------------------------------------------------------------
' The string representation of the given subject status
'----------------------------------------------------------------------------

    Select Case nBusyStatus
    Case eBusyStatus.SubjectBusy
        GetSubjectStatusString = "Busy"
    Case eBusyStatus.SubjectNotBusy
        GetSubjectStatusString = "Not Busy"
    End Select

End Function

'-------------------------------------------------------------------------
Public Function GetCacheReportB() As Variant
'-------------------------------------------------------------------------
' Returns an array giving information about each existing Cache Entry
'-------------------------------------------------------------------------
Dim vReport As Variant
Dim i As Integer
Dim nCount As Integer
Dim oCSubject As CachedSubject

    nCount = 0
    
    If Not mcolSCMSubjects Is Nothing Then
        nCount = mcolSCMSubjects.Count
        If nCount > 0 Then
            ReDim vReport(nCount + 1)
            ' vreport(0) will contain MaxArezzo value
            For i = 1 To nCount
                Set oCSubject = mcolSCMSubjects(i)
                ' Arrays start at 0, collections start at 1
                vReport(i) = CachedSubjectReport(i, oCSubject)
            Next i
        End If
    End If
    
    If nCount = 0 Then
        ' No cache entries
        ReDim vReport(0)
        vReport(0) = "No Cache Entries, "
    End If
    
    vReport(0) = vReport(0) & "MaxArezzoAllowed = " & gnMaxArezzoAllowed
    
    Set oCSubject = Nothing
    GetCacheReportB = vReport

End Function

'-------------------------------------------------------------------------
Private Function CachedSubjectReport(nEntry As Integer, oCSubject As CachedSubject) As String
'-------------------------------------------------------------------------
' Report on a single cache entry
' NCJ 15 Jan 04 - Changed separator from comma to pipe (cures regional settings problem)
' ic 18/05/2005 added database id
'-------------------------------------------------------------------------
Dim sReport As String

    sReport = ""
    sReport = sReport & "CacheEntry = " & nEntry
    sReport = sReport & "| Status = " & oCSubject.BusyStatusString
    sReport = sReport & "| Study = " & oCSubject.StudyId
    sReport = sReport & "| Site = " & oCSubject.Site
    sReport = sReport & "| SubjectID = " & oCSubject.SubjectId
    sReport = sReport & "| Timestamp = " & Format(CDate(oCSubject.TimeStamp), "dd/mm/yy, hh:mm:ss")
    sReport = sReport & "| CacheKey = " & oCSubject.MyKey
    sReport = sReport & "| CacheToken = " & oCSubject.ArezzoToken
    sReport = sReport & "| User = " & oCSubject.UserName
    sReport = sReport & "| DB = " & GetDatabaseId(oCSubject.DatabaseCnn)
    
    CachedSubjectReport = sReport
    
End Function

'-------------------------------------------------------------------------
Private Function GetDatabaseId(ByVal sDatabaseCnn As String) As String
'-------------------------------------------------------------------------
' Returns a database id extracted from a connection string
'-------------------------------------------------------------------------
Dim sDB As String
    
    On Error GoTo IgnoreError
    
    'sql server?
    sDB = Connection_Property(CONNECTION_DATABASE, sDatabaseCnn)
    If (sDB = "") Then
        'oracle
        sDB = Connection_Property(CONNECTION_USERID, sDatabaseCnn)
    End If
    
IgnoreError:
    GetDatabaseId = sDB
End Function

'-------------------------------------------------------------------------
Private Function NewKey() As String
'-------------------------------------------------------------------------
' Return the next key value for a CachedSubject object in our collection
'-------------------------------------------------------------------------

    NewKey = "Key" & mnKeyValue
    mnKeyValue = mnKeyValue + 1

End Function

'----------------------------------------------------------------------------------------'
Private Sub GetPrologSwitches(lTrialId As Long, sDBCon As String)
'----------------------------------------------------------------------------------------'
' NCJ 29 Jan 03
' Set up the Prolog switches the first time they're needed, then keep them the same
' Sets the value of msPrologSwitches
' Uses new clsArezzoMemory (not registry)
'----------------------------------------------------------------------------------------'
Dim oArezzoMemory As clsAREZZOMemory

    ' If not yet set up, go get them
    If msPrologSwitches = "" Then
        Set oArezzoMemory = New clsAREZZOMemory
        ' Read memory values from DB
        Call oArezzoMemory.Load(lTrialId, sDBCon)
        ' Get the memory settings, overriding from Settings file if appropriate
        msPrologSwitches = oArezzoMemory.AREZZOSwitches(True)
        Set oArezzoMemory = Nothing
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub ReportBusyCacheEntries()
'----------------------------------------------------------------------------------------'
' NCJ 29 May 03
' Report any cache entries that have been busy for more than 10 secs.
'----------------------------------------------------------------------------------------'
Dim dblTarget As Double
Dim oCS As CachedSubject
Dim sLog As String
Dim n As Integer
Dim sFileName As String
Dim i As Integer

    ' Ignore errors
    On Error Resume Next
    
    If Not TraceFlag Then Exit Sub
    If mcolSCMSubjects Is Nothing Then Exit Sub
    If mcolSCMSubjects.Count = 0 Then Exit Sub
    
    sLog = ""
    ' Subtract 10 secs. from current time
    dblTarget = CDbl(DateAdd("s", -10, Now))
    For i = 1 To mcolSCMSubjects.Count
        Set oCS = mcolSCMSubjects(i)
        ' Was it made busy more than 10 secs. ago?
        If oCS.BusyStatus = SubjectBusy And oCS.TimeStamp <= dblTarget Then
            ' We've got a suspect cache entry here
            sLog = sLog & vbCrLf & "*** SUSPECT BUSY CACHE ENTRY - "
            sLog = sLog & CachedSubjectReport(i, oCS)
        End If
    Next

    Set oCS = Nothing

    If sLog > "" Then
        sLog = Format(Now, "yyyy/dd/mm hh:mm:ss") & vbCrLf & sLog
        ' Write to special file
        n = FreeFile
        sFileName = App.Path & "\Temp\BusyLog.dat"
        Open sFileName For Append As n
    
        Print #n, sLog
    
        Close n
    End If
    
End Sub

'------------------------------------------------------------------------------'
Private Function TraceFlag() As Boolean
'------------------------------------------------------------------------------'
' ic 24/04/2003
' function returns trace on/off boolean
'------------------------------------------------------------------------------'
Dim sTrace As String

    Call InitialiseSettingsFile(True)
    sTrace = GetMACROSetting(MACRO_SETTING_TRACE, "false")
    TraceFlag = (LCase(sTrace) = "true")
    
End Function

'------------------------------------------------------------------------------'
Public Function LoadEformB(ByRef sSerialisedUser As String, ByVal lStudyId As Long, ByVal sSite As String, _
    ByVal lSubjectId As Long, ByVal sToken As String, ByVal lCRFPageTaskId As Long, ByRef sEFILockToken As String, _
    ByRef sVILockToken As String, ByRef bEFIUnavailable As Boolean, ByVal sDecimalPoint As String, _
    ByVal sThousandSeparator As String, Optional ByVal vAlerts As Variant, Optional ByVal vErrors As Variant, _
    Optional ByVal bAutoNext As Boolean = False) As String
'------------------------------------------------------------------------------'
' ic 18/08/2003
' function finds corresponding subject object in cache, passes it to IOEform
' object to be used to generate eform html
' revisions
' ic 20/12/2004 bug 2395 - pass the sSerialisedUser byref so that lastused eform gets set and passed back
' ic 18/05/2005 issue 2560, pass database connection string
'------------------------------------------------------------------------------'
Dim oUser As MACROUser
Dim oCSubject As CachedSubject
Dim sRtn As String
Dim oIO As MACROIOEform30.WWWIOEform
Dim bTrace As Boolean

    On Error GoTo handler

    Set oUser = New MACROUser
    Call oUser.SetState(CStr(sSerialisedUser))
    bTrace = TraceFlag()
    
    ' Look for Cache entry with this token, with status Busy
    Set oCSubject = GetCachedSubject(oUser.CurrentDBConString, lStudyId, sSite, lSubjectId, sToken, SubjectBusy)
    ' We ought to have one, but do nothing if we didn't find it
    If Not oCSubject Is Nothing Then
        If oCSubject.BusyStatus = SubjectBusy Then
            Set oIO = New MACROIOEform30.WWWIOEform
            
            LoadEformB = oIO.LoadEform(oUser, oCSubject.Subject, sSite, lCRFPageTaskId, sEFILockToken, _
                sVILockToken, bEFIUnavailable, sDecimalPoint, sThousandSeparator, vAlerts, vErrors, bAutoNext)
        
            sSerialisedUser = oUser.GetState(False)
            Set oIO = Nothing
        End If
        Set oCSubject = Nothing
    End If
    Set oUser = Nothing
    Exit Function
    
handler:
    Err.Raise Err.Number, , Err.Description & "|" & "modSCM.LoadEformB( " & lStudyId & "," & sSite & "," & lSubjectId & " )"
End Function

'------------------------------------------------------------------------------'
Public Function SaveEformB(ByVal sSerialisedUser As String, ByVal lStudyId As Long, ByVal sSiteCode As String, _
    ByVal lSubjectId As Long, ByVal sToken As String, ByVal sCRFPageTaskId As String, ByVal sForm As String, _
    ByRef sEFILockToken As String, ByRef sVILockToken As String, ByVal bVReadOnly As Boolean, _
    ByVal bEReadOnly As Boolean, ByVal sLabCode As String, ByVal sDecimalPoint As String, _
    ByVal sThousandSeparator As String, ByRef sRegister As String, ByVal sLocalDate As String, _
    Optional ByVal nTimezoneOffset As Integer = 0) As Variant
'------------------------------------------------------------------------------'
' ic 18/08/2003
' function finds corresponding subject object in cache, passes it to IOEform
' object to be used to save eform
' ic 18/05/2005 issue 2560, pass database connection string
'------------------------------------------------------------------------------'
Dim oUser As MACROUser
Dim oCSubject As CachedSubject
Dim sRtn As String
Dim oIO As MACROIOEform30.WWWIOEform
Dim bTrace As Boolean

    On Error GoTo handler

    Set oUser = New MACROUser
    Call oUser.SetState(CStr(sSerialisedUser))
    bTrace = TraceFlag()
    
    ' Look for Cache entry with this token, with status Busy
    Set oCSubject = GetCachedSubject(oUser.CurrentDBConString, lStudyId, sSiteCode, lSubjectId, sToken, SubjectBusy)
    ' We ought to have one, but do nothing if we didn't find it
    If Not oCSubject Is Nothing Then
        If oCSubject.BusyStatus = SubjectBusy Then
            Set oIO = New MACROIOEform30.WWWIOEform
            
            SaveEformB = oIO.SaveEform(oUser, oCSubject.Subject, sCRFPageTaskId, sForm, sEFILockToken, _
                sVILockToken, bVReadOnly, bEReadOnly, sLabCode, sDecimalPoint, sThousandSeparator, sRegister, _
                sLocalDate, nTimezoneOffset)
        
            Set oIO = Nothing
        End If
        Set oCSubject = Nothing
    End If
    Set oUser = Nothing
    Exit Function
    
handler:
    Err.Raise Err.Number, , Err.Description & "|" & "modSCM.SaveEformB( " & lStudyId & "," & sSiteCode & "," & lSubjectId & " )"
End Function

'ic 16/03/2004 remove conditional compilation
'#If ORAMA = 1 Then
'------------------------------------------------------------------------------'
Public Function CheckArezzoEventsB(ByVal sDatabaseCnn As String, ByVal lStudyId As Long, ByVal sSiteCode As String, ByVal lSubjectId As Long, _
    ByVal sToken As String, ByVal sDatabase As String, ByVal sEformPageTaskId As String, ByVal sNext As String, _
    ByRef bArezzoEvents As Boolean) As String
'------------------------------------------------------------------------------'
' ic 12/09/2003
' function finds corresponding subject object in cache, passes it to IOEform
' object to be used to check arezzo events
' ic 18/05/2005 issue 2560, pass database connection string
'------------------------------------------------------------------------------'
Dim oCSubject As CachedSubject
Dim sRtn As String
Dim oIO As MACROIOEform30.WWWIOEform
Dim bTrace As Boolean

    On Error GoTo handler
    
    ' Look for Cache entry with this token, with status Busy
    Set oCSubject = GetCachedSubject(sDatabaseCnn, lStudyId, sSiteCode, lSubjectId, sToken, SubjectBusy)
    ' We ought to have one, but do nothing if we didn't find it
    If Not oCSubject Is Nothing Then
        If oCSubject.BusyStatus = SubjectBusy Then
            Set oIO = New MACROIOEform30.WWWIOEform
            
            CheckArezzoEventsB = oIO.CheckArezzoEventsA(oCSubject.Subject, sDatabase, sEformPageTaskId, sNext, bArezzoEvents)
        
            Set oIO = Nothing
        End If
        Set oCSubject = Nothing
    End If
    Exit Function
    
handler:
    Err.Raise Err.Number, , Err.Description & "|" & "modSCM.CheckArezzoEventsB( " & lStudyId & "," & sSiteCode & "," & lSubjectId & " )"
End Function

'------------------------------------------------------------------------------'
Public Sub SaveArezzoEventsB(ByVal sDatabaseCnn As String, ByVal lStudyId As Long, ByVal sSiteCode As String, ByVal lSubjectId As Long, _
    ByVal sToken As String, ByVal sForm As String)
'------------------------------------------------------------------------------'
' ic 12/09/2003
' function finds corresponding subject object in cache, passes it to IOEform
' object to be used to save arezzo events
' ic 18/05/2005 issue 2560, pass database connection string
'------------------------------------------------------------------------------'
Dim oCSubject As CachedSubject
Dim oIO As MACROIOEform30.WWWIOEform

    On Error GoTo handler

    ' Look for Cache entry with this token, with status Busy
    Set oCSubject = GetCachedSubject(sDatabaseCnn, lStudyId, sSiteCode, lSubjectId, sToken, SubjectBusy)
    ' We ought to have one, but do nothing if we didn't find it
    If Not oCSubject Is Nothing Then
        If oCSubject.BusyStatus = SubjectBusy Then
            Set oIO = New MACROIOEform30.WWWIOEform
            
            Call oIO.SaveArezzoEventsA(oCSubject.Subject, sForm)
        
            Set oIO = Nothing
        End If
        Set oCSubject = Nothing
    End If
    Exit Sub
    
handler:
    Err.Raise Err.Number, , Err.Description & "|" & "modSCM.SaveArezzoEventsB( " & lStudyId & "," & sSiteCode & "," & lSubjectId & " )"
End Sub
'#End If
