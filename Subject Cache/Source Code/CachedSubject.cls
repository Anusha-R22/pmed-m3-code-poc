VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CachedSubject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------
'   File:       CachedSubject.cls
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   Author:     Nicky Johns, July 2002
'   Purpose:    A MACRO subject cached by the Cache Manager
'-----------------------------------------------------------------------------
' Revisions:
'   NCJ 18-22 July - Initial development
'   NCJ 23 Jul 02 - Added Reload function
'   MACRO 3.0
'   NCJ 17 Sept 02 - Updated for new MACRO 3.0 locking model
'   NCJ 24 Jan 03 - Added sCountry parameter to NewSubject
'   NCJ 27 May 03 - Tightened up on error handling (new mbSubjectLoaded variable)
'   NCJ 26 Nov 03 - Added extra WriteLog after starting Prolog in Init
'   ic 10/11/2004 bug 2439, call SetUserProperties() in Reload()
'   ic 18/05/2005 issue 2560, added database connection string 'Get'
'----------------------------------------------------------------------------

Option Explicit

Private moArezzo As Arezzo_DM
Private moStudyDef As StudyDefRO
Private menBusyStatus As eBusyStatus
Private mdblTimeStamp As Double
'Private mlDBToken As Long
Private msKey As String
Private msArezzoToken As String
Private msDatabaseCnn As String
Private mnUpdateMode As eUIUpdateMode
Private msUser As String
Private mbSubjectLoaded As Boolean  ' Whether we have a successfully loaded subject

'----------------------------------------------------------------------------
Public Sub Init(ByVal sKey As String, _
            ByVal sTempDir As String, ByVal sPrologSwitches As String)
'----------------------------------------------------------------------------
' Initialise this Cache object
' NB This MUST be done before its first use
' sKey is its unique key (used when adding to the collection)
' sTempDir is the temp. directory for Arezzo to use,
' and sPrologSwitches is the Prolog memory settings.
' Need to create a new Arezzo and a new StudyDef object
'----------------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    ' Create a new Arezzo object
    ' This Arezzo/Prolog instance stays for the lifetime of this Cache object
    Set moArezzo = New Arezzo_DM
    ' Set it going (i.e. start Prolog)
    Call moArezzo.Init(sTempDir, sPrologSwitches)
    
    Call WriteLog("Started new Prolog instance for Cache object " & sKey)
    
    ' This StudyDefRO instance also persists for the lifetime of this Cache object
    ' (although it may contain different studies)
    Set moStudyDef = New StudyDefRO

    ' No Arezzo token yet
    msArezzoToken = ""
    
    mdblTimeStamp = CDbl(Now)
    menBusyStatus = SubjectNotBusy
    msKey = sKey
    mbSubjectLoaded = False
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|CachedSubject.Init"

End Sub

'----------------------------------------------------------------------------
Public Function NewSubject(lStudyId As Long, sSite As String, _
                sUserName As String, sCountry As String, _
                sDatabaseCnn As String, _
                ByVal sUserNameFull As String, _
                ByVal sUserRole As String) As StudySubject
'----------------------------------------------------------------------------
' Create a new subject for the specified Study and Site
' Returns created StudySubject object (may be Nothing if unsuccessful)
' NCJ 24 Jan 03 - Added sCountry parameter
' NCJ 6 May 03 - Added new params. sUserNameFull and sUserRole
'----------------------------------------------------------------------------
Dim oSubject As StudySubject

    On Error GoTo ErrLabel
    
    ' Tidy up before creating a new one
    Call moStudyDef.RemoveSubject
    
    ' Indicate subject not loaded
    mbSubjectLoaded = False
    
    Call moStudyDef.Load(sDatabaseCnn, lStudyId, 1, moArezzo)

    Set oSubject = moStudyDef.NewSubject(sSite, sUserName, sCountry, sUserNameFull, sUserRole)

    ' Store the database where we loaded from
    msDatabaseCnn = sDatabaseCnn
    
    ' NCJ 27 May 03 - Check we've got a valid subject
    If Not oSubject Is Nothing Then
        If Not oSubject.CouldNotLoad Then
            ' Indicate subject loaded
            mbSubjectLoaded = True
            
            ' Assume this is Read-Write
            mnUpdateMode = Read_Write
            
            ' Store our Arezzo (cache) token
            msArezzoToken = moStudyDef.Subject.CacheToken
            
            ' Set the status to be Busy
            menBusyStatus = SubjectBusy
            mdblTimeStamp = CDbl(Now)
        End If
    End If
    
    Set NewSubject = oSubject
    Set oSubject = Nothing

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "| CachedSubject.NewSubject( " _
                    & lStudyId & ", " & sSite & " )"

End Function

'----------------------------------------------------------------------------
Public Function Load(lStudyId As Long, sSite As String, lSubjectId As Long, _
                sUserName As String, nUpdateMode As Integer, _
                sDatabaseCnn As String, _
                ByVal sUserNameFull As String, _
                ByVal sUserRole As String) As StudySubject
'----------------------------------------------------------------------------
' Load (or reload) the specified subject
' Returns created StudySubject object
' NCJ 6 May 03 - Added new params. sUserNameFull and sUserRole
'----------------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    ' Tidy up before creating a new one
    Call moStudyDef.RemoveSubject
    mbSubjectLoaded = False
    
    Call moStudyDef.Load(sDatabaseCnn, lStudyId, 1, moArezzo)
    Set Load = moStudyDef.LoadSubject(sSite, lSubjectId, sUserName, nUpdateMode, sUserNameFull, sUserRole, False)
    ' NB The Higher level routines will check the result of the load
    
    ' Store the database where we loaded from
    msDatabaseCnn = sDatabaseCnn
    
    ' NCJ 27 May 03 - Also check here
    If Not moStudyDef.Subject Is Nothing Then
        If Not moStudyDef.Subject.CouldNotLoad Then
            mbSubjectLoaded = True
            
            ' Store the update mode
            mnUpdateMode = nUpdateMode
            
            ' Store our Arezzo token
            msArezzoToken = moStudyDef.Subject.CacheToken
            
            ' Set the status to be Busy
            menBusyStatus = SubjectBusy
            mdblTimeStamp = CDbl(Now)
        End If
    End If

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "| CachedSubject.Load( " _
                    & lStudyId & ", " & sSite & ", " & lSubjectId & " )"

End Function

'----------------------------------------------------------------------------
Public Function Reload(sUserName As String, nUpdateMode As Integer, _
                        ByVal sUserNameFull As String, _
                        ByVal sUserRole As String) As StudySubject
'----------------------------------------------------------------------------
' Reload the existing subject if necessary for the given User and UpdateMode
' i.e. if it's valid we don't need to reload,
' but if it's invalid we do need to reload
' Returns created StudySubject object
' NCJ 6 May 03 - Added new params. sUserNameFull and sUserRole
' ic 10/11/2004 bug 2439, call SetUserProperties()
'----------------------------------------------------------------------------

    On Error GoTo ErrLabel
    
    If NeedToReload(nUpdateMode) Then
        ' Just reload everything
        Set Reload = Me.Load(Me.StudyId, Me.Site, Me.SubjectId, sUserName, nUpdateMode, msDatabaseCnn, _
                            sUserNameFull, sUserRole)
    Else
        ' Ensure we have the right user
        moStudyDef.Subject.UserName = sUserName
        Call moStudyDef.Subject.SetUserProperties(sUserNameFull, sUserRole)
        ' Set the status to be Busy
        menBusyStatus = SubjectBusy
        mdblTimeStamp = CDbl(Now)
        Set Reload = Me.Subject
    End If

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "| CachedSubject.Reload( " _
                    & sUserName & ", " & nUpdateMode & " )"

End Function

'----------------------------------------------------------------------------
Private Function NeedToReload(nUpdateMode As Integer) As Boolean
'----------------------------------------------------------------------------
' Do we need to reload this subject?
'----------------------------------------------------------------------------

    NeedToReload = True
    
    ' Check whether we were successfully loaded to start with
    ' (Check all these just to be on the safe side!)
    If Not mbSubjectLoaded Then Exit Function
    If moStudyDef.Subject Is Nothing Then Exit Function
    If moStudyDef.Subject.CouldNotLoad Then Exit Function
    
    ' Reload if Update mode is different
    If nUpdateMode <> mnUpdateMode Then Exit Function
    
    ' Reload if invalid
    If IsArezzoInvalid(msArezzoToken, Me.StudyId, Me.Site, Me.SubjectId, msDatabaseCnn) Then Exit Function
    
    ' If we get here, we're OK and don't need to reload
    NeedToReload = False
    
End Function

'----------------------------------------------------------------------------
Public Sub ReleaseSubject()
'----------------------------------------------------------------------------
' Say that we're no longer using this subject
'----------------------------------------------------------------------------

    menBusyStatus = SubjectNotBusy

End Sub

'----------------------------------------------------------------------------
Public Property Get Subject() As StudySubject
'----------------------------------------------------------------------------
' The subject
'----------------------------------------------------------------------------

    Set Subject = moStudyDef.Subject

End Property

'----------------------------------------------------------------------------
Public Property Get TimeStamp() As Double
'----------------------------------------------------------------------------
' The timestamp at which this object was last used
'----------------------------------------------------------------------------

    TimeStamp = mdblTimeStamp

End Property

'----------------------------------------------------------------------------
Public Property Get BusyStatus() As eBusyStatus
'----------------------------------------------------------------------------
' The status of this object as an integer
'----------------------------------------------------------------------------

    BusyStatus = menBusyStatus

End Property

'----------------------------------------------------------------------------
Public Property Get BusyStatusString() As String
'----------------------------------------------------------------------------
' The status of this object as a string
'----------------------------------------------------------------------------

    BusyStatusString = GetSubjectStatusString(menBusyStatus)

End Property

'----------------------------------------------------------------------------
Public Property Get ArezzoToken() As String
'----------------------------------------------------------------------------
' The DB Token for the currently loaded study/site/subject
'----------------------------------------------------------------------------
    
    ArezzoToken = msArezzoToken
    
End Property

'----------------------------------------------------------------------------
Public Property Get StudyId() As Long
'----------------------------------------------------------------------------
' The ID of the loaded study
'----------------------------------------------------------------------------
    
    StudyId = moStudyDef.StudyId

End Property

'----------------------------------------------------------------------------
Public Property Get SubjectId() As Long
'----------------------------------------------------------------------------
' The ID of the loaded subject (if any)
' Returns -1 if subject not loaded
'----------------------------------------------------------------------------
    
    If mbSubjectLoaded And (Not moStudyDef.Subject Is Nothing) Then
        SubjectId = moStudyDef.Subject.PersonID
    Else
        SubjectId = -1
    End If
    
End Property

'----------------------------------------------------------------------------
Public Property Get Site() As String
'----------------------------------------------------------------------------
' The Site of the loaded subject (if any)
'----------------------------------------------------------------------------
    
    If mbSubjectLoaded And (Not moStudyDef.Subject Is Nothing) Then
        Site = moStudyDef.Subject.Site
    Else
        Site = ""
    End If
    
End Property

'----------------------------------------------------------------------------
Public Property Get DatabaseCnn() As String
'----------------------------------------------------------------------------
' The database connection string of the loaded subject (if any)
'----------------------------------------------------------------------------
    
    DatabaseCnn = msDatabaseCnn
    
End Property

'----------------------------------------------------------------------------
Public Property Get UserName() As String
'----------------------------------------------------------------------------
' The User name of the loaded subject (if any)
'----------------------------------------------------------------------------
    
    If mbSubjectLoaded And (Not moStudyDef.Subject Is Nothing) Then
        UserName = moStudyDef.Subject.UserName
    Else
        UserName = ""
    End If

End Property

'----------------------------------------------------------------------------
Public Property Get MyKey() As String
'----------------------------------------------------------------------------
' The key with which we were added to the collection
'----------------------------------------------------------------------------

    MyKey = msKey

End Property

'----------------------------------------------------------------------------
Private Sub Class_Initialize()
'----------------------------------------------------------------------------
' Class is being created
'----------------------------------------------------------------------------

    ' Do nothing here, but we expect Init to be called instead

End Sub

'----------------------------------------------------------------------------
Private Sub Class_Terminate()
'----------------------------------------------------------------------------
' Class is being destroyed so tidy everything up
'----------------------------------------------------------------------------

    Call Terminate

End Sub

'----------------------------------------------------------------------------
Public Sub Terminate()
'----------------------------------------------------------------------------
' This is to tidy up before destroying this object
'----------------------------------------------------------------------------

    If Not moStudyDef Is Nothing Then
        ' Tidy up the study definition object
        Call moStudyDef.RemoveSubject
        Call moStudyDef.Terminate
        Set moStudyDef = Nothing
    End If
    
    If Not moArezzo Is Nothing Then
        ' Shut down its Prolog
        Call moArezzo.Finish
        Set moArezzo = Nothing
    End If
    
End Sub


