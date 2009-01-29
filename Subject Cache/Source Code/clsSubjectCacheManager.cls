VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubjectCacheManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------
'   File:       clsSubjectCacheManager.cls
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   Author:     Nicky Johns, July 2002 (based on original by Zulfi)
'   Purpose:    Manages the Arezzo Subject Cache objects
'               NB All Calls in here go through to modSubjectCacheManager
'-----------------------------------------------------------------------------
' Revisions:
'    NCJ 17 Sept 02 - Changes for new MACRO 3.0 Locking model
'   NCJ 24 Jan 03 - Added sCountry parameter to NewSubject
'   ic 18/08/2003 added LoadEformA() and SaveEformA() to improve performance on certain machines
'   ic 15/09/2003 added CheckArezzoEventsA() and SaveArezzoEventsA() to improve performance on certain machines
'   ic 16/03/2004 remove conditional ORAMA compilation
'   ic 20/12/2004 bug 2395 - pass the sSerialisedUser byref so that lastused eform gets set and passed back
'   ic 18/05/2005 issue 2560, pass database connection for locating cached subjects
'----------------------------------------------------------------------------

Option Explicit


''-------------------------------------------------------------------------
'Public Sub MarkArezzoInvalid(ByVal lStudyId As Long, _
'                            ByVal sSite As String, _
'                            ByVal lSubjectId As Long, _
'                            ByVal sConnString As String)
''-------------------------------------------------------------------------
'' Mark the specified subject as invalid
''-------------------------------------------------------------------------
'
'    Call MarkAsInvalidB(lStudyId, sSite, lSubjectId, sConnString)
'
'End Sub

'--------------------------------------------------------------------------------------------------
Public Function CreateNewSubject(ByVal sUserName As String, _
                                ByVal sDatabaseCnn As String, _
                                ByVal lStudyId As Long, _
                                ByVal sSite As String, _
                                ByVal sCountry As String, _
                                ByVal sUserNameFull As String, _
                                ByVal sUserRole As String) As Variant
'--------------------------------------------------------------------------------------------------
' Create new subject and return array containing:
'   Result code (0 if OK)
'   New StudySubject
'   ArezzoToken for the subject
' NCJ 24 Jan 03 - Added sCountry parameter
' NCJ 6 May 03 - Added new params. sUserNameFull and sUserRole
'--------------------------------------------------------------------------------------------------

    CreateNewSubject = CreateNewSubjectA(sUserName, sDatabaseCnn, lStudyId, sSite, sCountry, _
                            sUserNameFull, sUserRole)
 
 End Function

'--------------------------------------------------------------------------------
Public Function LoadSubject(ByVal sDatabaseCnn As String, _
                        ByVal lStudyId As Long, _
                        ByVal sSite As String, _
                        ByVal lSubjectId As Long, _
                        ByVal sUserName As String, _
                        ByVal nUpdateMode As Integer, _
                        ByVal sUserNameFull As String, _
                        ByVal sUserRole As String) As Variant
'--------------------------------------------------------------------------------
'The need to pass Connection string rather than database name will help Arezzo
'Cache to avoid reading through the registry. This must be done at IO level and
'pass it through Arezzo Cache Manager
' NCJ 6 May 03 - Added new params. sUserNameFull and sUserRole
'--------------------------------------------------------------------------------
    
    LoadSubject = LoadSubjectA(sDatabaseCnn, lStudyId, sSite, lSubjectId, sUserName, nUpdateMode, _
                            sUserNameFull, sUserRole)
    
End Function

'-------------------------------------------------------------------------
Public Function ReleaseSubject(ByVal sDatabaseCnn As String, ByVal lStudyId As Long, _
                            ByVal sSite As String, _
                            ByVal lSubjectId As Long, _
                            ByVal sCacheToken As String) As Boolean
'-------------------------------------------------------------------------
'when a request for release arrives, we change the subject state to NotBusy
'so that a new request for same subject can be served immediately
'ic 18/05/2005 issue 2560, pass database connection for locating cached subjects
'-------------------------------------------------------------------------

    ReleaseSubject = ReleaseSubjectB(sDatabaseCnn, lStudyId, sSite, lSubjectId, sCacheToken)

End Function

'--------------------------------------------------------------------------------
Public Property Get MaxArezzoAllowed() As Integer
'--------------------------------------------------------------------------------
'Retrieve the maximum number of Arezzo allowed to load by this server
'--------------------------------------------------------------------------------
    
    MaxArezzoAllowed = gnMaxArezzoAllowed

End Property

'--------------------------------------------------------------------------------
Public Property Let MaxArezzoAllowed(ByVal nMaxArezzoAllowed As Integer)
'--------------------------------------------------------------------------------
'set the maximum number of Arezzo allowed to load by this server
'--------------------------------------------------------------------------------

    ' If no change then do nothing
    If nMaxArezzoAllowed = gnMaxArezzoAllowed Then Exit Property
    
    If nMaxArezzoAllowed > gnMaxArezzoAllowed Then
        ' We just increase the value
        gnMaxArezzoAllowed = nMaxArezzoAllowed
    Else
        ' Here we need to kill off the "surplus" Arezzos
        ' (the new value is less than the existing value)
        ' but we only do as many as we can
        ' Reset max to how many are left
        gnMaxArezzoAllowed = ReduceCacheObjects(nMaxArezzoAllowed)
    End If
     
End Property

'-------------------------------------------------------------------------
Public Function ClearArezzoTokenTable(ByVal sConnString As String)
'-------------------------------------------------------------------------
'clear ArezzoToken table if user has decided to shut down the server
'-------------------------------------------------------------------------
    
'    Call ClearArezzoTokenTableA(sConnString)

End Function

'-------------------------------------------------------------------------
Public Sub Terminate()
'-------------------------------------------------------------------------
' Tidy up before we go home
' NB This will close down all objects and terminate all Prolog instances
'-------------------------------------------------------------------------

    Call CloseSCM

End Sub

'-------------------------------------------------------------------------
Private Sub Class_Initialize()
'-------------------------------------------------------------------------
' We're being created so initialise the necessaries
'-------------------------------------------------------------------------

    Call InitSCM

End Sub

'-------------------------------------------------------------------------
Private Sub Class_Terminate()
'-------------------------------------------------------------------------
' We're being destroyed - make sure we close everything down
'-------------------------------------------------------------------------

    ' This is no good when running from a VB exe
    ' because it shuts down all our cache objects!!!
'    Call Terminate

End Sub

'-------------------------------------------------------------------------
Public Function GetCacheReport() As Variant
'-------------------------------------------------------------------------
' Returns an array giving information about each existing Cache Entry
'-------------------------------------------------------------------------

    GetCacheReport = GetCacheReportB()
    
End Function

'-------------------------------------------------------------------------
Public Function LoadEformA(ByRef sSerialisedUser As String, ByVal lStudyId As Long, ByVal sSite As String, _
    ByVal lSubjectId As Long, ByVal sToken As String, ByVal lCRFPageTaskId As Long, ByRef sEFILockToken As String, _
    ByRef sVILockToken As String, ByRef bEFIUnavailable As Boolean, ByVal sDecimalPoint As String, _
    ByVal sThousandSeparator As String, Optional ByVal vAlerts As Variant, Optional ByVal vErrors As Variant, _
    Optional ByVal bAutoNext As Boolean = False) As String
'-------------------------------------------------------------------------
' ic 18/08/2003
' function wrapper for LoadEformB
' revisions
' ic 20/12/2004 bug 2395 - pass the sSerialisedUser byref so that lastused eform gets set and passed back
'-------------------------------------------------------------------------
    LoadEformA = LoadEformB(sSerialisedUser, lStudyId, sSite, lSubjectId, sToken, lCRFPageTaskId, sEFILockToken, _
        sVILockToken, bEFIUnavailable, sDecimalPoint, sThousandSeparator, vAlerts, vErrors, bAutoNext)
End Function

'------------------------------------------------------------------------------'
Public Function SaveEformA(ByVal sSerialisedUser As String, ByVal lStudyId As Long, ByVal sSiteCode As String, _
    ByVal lSubjectId As Long, ByVal sToken As String, ByVal sCRFPageTaskId As String, ByVal sForm As String, _
    ByRef sEFILockToken As String, ByRef sVILockToken As String, ByVal bVReadOnly As Boolean, _
    ByVal bEReadOnly As Boolean, ByVal sLabCode As String, ByVal sDecimalPoint As String, _
    ByVal sThousandSeparator As String, ByRef sRegister As String, ByVal sLocalDate As String, _
    Optional ByVal nTimezoneOffset As Integer = 0) As Variant
'------------------------------------------------------------------------------'
' ic 18/08/2003
' function wrapper for SaveEformB
'------------------------------------------------------------------------------'
    SaveEformA = SaveEformB(sSerialisedUser, lStudyId, sSiteCode, lSubjectId, sToken, sCRFPageTaskId, sForm, _
        sEFILockToken, sVILockToken, bVReadOnly, bEReadOnly, sLabCode, sDecimalPoint, sThousandSeparator, _
        sRegister, sLocalDate, nTimezoneOffset)
End Function

'ic 16/03/2004 remove conditional compilation
'#If ORAMA = 1 Then
'------------------------------------------------------------------------------'
Public Function CheckArezzoEventsA(ByVal sDatabaseCnn As String, ByVal lStudyId As Long, ByVal sSite As String, ByVal lSubjectId As Long, _
    ByVal sToken As String, ByVal sDatabase As String, ByVal sEformPageTaskId As String, ByVal sNext As String, _
    ByRef bArezzoEvents As Boolean) As String
'------------------------------------------------------------------------------'
' ic 15/09/2003
' ic 18/05/2005 issue 2560, pass database connection for locating cached subjects
'------------------------------------------------------------------------------'
    CheckArezzoEventsA = CheckArezzoEventsB(sDatabaseCnn, lStudyId, sSite, lSubjectId, sToken, sDatabase, sEformPageTaskId, _
        sNext, bArezzoEvents)
End Function

'------------------------------------------------------------------------------'
Public Sub SaveArezzoEventsA(ByVal sDatabaseCnn As String, ByVal lStudyId As Long, ByVal sSite As String, ByVal lSubjectId As Long, _
    ByVal sToken As String, ByVal sForm As String)
'------------------------------------------------------------------------------'
'   12/09/2003
'   function wrapper for SaveArezzoEventsB
'   ic 18/05/2005 issue 2560, pass database connection for locating cached subjects
'------------------------------------------------------------------------------'
    Call SaveArezzoEventsB(sDatabaseCnn, lStudyId, sSite, lSubjectId, sToken, sForm)
End Sub
'#End If
