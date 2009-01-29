VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubjectCacheInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------
'   File:       clsSubjectCacheInterface.cls
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   Author:     Nicky Johns, July 2002 (based on original by Zulfi)
'   Purpose:    Interface to Subject Cache Manager
'-----------------------------------------------------------------------------
' Revisions:
'   NCJ 19 Sept 02 - Changes for MACRO 3.0
'   NCJ 24 Jan 03 - Added sCountry parameter to CreateNewSubject
'   ic 18/08/2003 added LoadEform() and SaveEform() to improve performance on certain machines
'   ic 15/09/2003 added CheckArezzoEvents() and SaveArezzoEvents() to improve performance on certain machines
'   ic 16/03/2004 remove conditional ORAMA compilation
'   ic 20/12/2004 bug 2395 - pass the sSerialisedUser byref so that lastused eform gets set and passed back
'   ic 18/05/2005 issue 2560, pass database connection for locating cached subjects
'----------------------------------------------------------------------------

Option Explicit

' How long we keep trying to achieve our goal in seconds (2 minutes)
Private Const msglPAUSE_TIME As Single = 120

' How long we wait between loops when trying (in milliseconds)
Private Const mlSLEEP_TIME As Long = 1000

'--------------------------------------------------------------------------------
Private Sub Class_Initialize()
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------

End Sub

'--------------------------------------------------------------------------------
Private Sub Class_Terminate()
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------

End Sub

'--------------------------------------------------------------------------------------------------
Public Function CreateNewSubject(ByVal sDatabaseCnn As String, _
                                 ByVal lStudyId As Long, _
                                 ByVal sSite As String, _
                                 ByVal sUserName As String, _
                                 ByVal sCountry As String, _
                                 ByVal sUserNameFull As String, _
                                 ByVal sUserRole As String) As Variant
'--------------------------------------------------------------------------------------------------
' Creates a new subject
' Returns an array containing Result Code, Study Subject and Token string
' NCJ 24 Jan 03 - Added sCountry parameter
' NCJ 6 May 03 - Added new params. sUserNameFull and sUserRole
'--------------------------------------------------------------------------------------------------
Dim vResult() As Variant
Dim oACM As clsSubjectCacheManager
Dim sglStartTime

    On Error GoTo ErrLabel
        
    sglStartTime = Timer
    
    ' Keep trying until we get one
    Do While Timer < sglStartTime + msglPAUSE_TIME
        
        ' NB We have to create a new ACM each loop iteration
        ' to allow others the chance of getting in between loops
        Set oACM = New clsSubjectCacheManager
            
        'call CreateNewSubject method to create a new subject
        vResult = oACM.CreateNewSubject(sUserName, sDatabaseCnn, _
                                lStudyId, sSite, sCountry, _
                                sUserNameFull, sUserRole)
                
        ' if we have got a subject, let's go out of this loop
        Select Case vResult(0)
        Case eCacheLoadResult.clrOK
            ' Success
            WriteLog "CreateNewSubject::Created " & " " & lStudyId & "," & sSite
            Exit Do
        Case eCacheLoadResult.clrNoCacheEntries, eCacheLoadResult.clrStudyLocked
            ' Write out failure reason
            WriteLog "CreateNewSubject::FAILED " & vResult(2) & " " & lStudyId & "," & sSite
        End Select
        
        Set oACM = Nothing
        
        ' Wait a bit before looping again
        Sleep mlSLEEP_TIME
        
    Loop
    
    CreateNewSubject = vResult
    
    Set oACM = Nothing

    Exit Function
    
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & _
        "clsSubjectCacheInterface.CreateNewSubject( " & lStudyId & ", " & sSite & " )"

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
' Load the specified subject
' Returns an array containing Result Code, Study Subject and Token string
' NCJ 6 May 03 - Added new params. sUserNameFull and sUserRole
'--------------------------------------------------------------------------------
Dim oACM As clsSubjectCacheManager
Dim sglStartTime As Single
Dim vResult() As Variant

    On Error GoTo ErrLabel
    
    ' We return an array of three results
    ReDim vResult(2)
    
    sglStartTime = Timer
    
    Do While Timer < sglStartTime + msglPAUSE_TIME
        
        ' NB We have to create a new ACM each loop iteration
        ' to allow others the chance of getting in between loops
        Set oACM = New clsSubjectCacheManager
        
        'call GetSubject method to load a subject
         vResult = oACM.LoadSubject(sDatabaseCnn, lStudyId, sSite, lSubjectId, sUserName, nUpdateMode, _
                            sUserNameFull, sUserRole)
        
        ' if we have got a subject, let's go out of this loop
        Select Case vResult(0)
        Case eCacheLoadResult.clrOK
            ' Success
            WriteLog "LoadSubject::Loaded " & GetToken(lStudyId, sSite, lSubjectId)
            Exit Do
        Case eCacheLoadResult.clrNoCacheEntries, eCacheLoadResult.clrStudyLocked
            WriteLog "LoadSubject::FAILED " & vResult(2) & " " & GetToken(lStudyId, sSite, lSubjectId)
        End Select
        
        Set oACM = Nothing
    
        ' Wait a bit before looping again
        Sleep mlSLEEP_TIME
    Loop
    
    ' Return the result
    LoadSubject = vResult
    
    Set oACM = Nothing
    
    Exit Function

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "clsSubjectCacheInterface.LoadSubject( " _
                    & lStudyId & ", " & sSite & ", " & lSubjectId & " )"
    
End Function

'--------------------------------------------------------------------------------
Public Sub ReleaseSubject(ByVal sDatabaseCnn As String, ByVal lStudyId As Long, _
                            ByVal sSite As String, _
                            ByVal lSubjectId As Long, _
                            ByVal sCacheToken As String)
'--------------------------------------------------------------------------------
' Release a subject specified by lStudyId, sSite and lSubjectId and sCacheToken
' i.e. mark it as not busy
' ic 18/05/2005 issue 2560, pass database connection for locating cached subjects
'--------------------------------------------------------------------------------
Dim oACM As clsSubjectCacheManager
        
    On Error GoTo ErrLabel
    
    Set oACM = New clsSubjectCacheManager
    
    ' Log success or otherwise
    If oACM.ReleaseSubject(sDatabaseCnn, lStudyId, sSite, lSubjectId, sCacheToken) Then
        WriteLog "ReleaseSubject::Released " & GetToken(lStudyId, sSite, lSubjectId)
    Else
        WriteLog "ReleaseSubject::FAILED " & GetToken(lStudyId, sSite, lSubjectId)
    End If
                 
    Set oACM = Nothing
        
    Exit Sub
    
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & _
            "clsSubjectCacheInterface.ReleaseSubject( " & lStudyId & "," & sSite & "," & lSubjectId & " )"
    
End Sub

'--------------------------------------------------------------------------------------------------
Public Sub ShutDownCacheManager()
'--------------------------------------------------------------------------------------------------
' This shuts down the Cache Manager
' NB All Arezzos will be shut down and all cached subjects lost!
'--------------------------------------------------------------------------------------------------
Dim oACM As clsSubjectCacheManager

    On Error GoTo ErrLabel
        
    Set oACM = New clsSubjectCacheManager
    Call oACM.Terminate
    WriteLog "Cache Manager Shut Down"
    Set oACM = Nothing
    
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & _
        "clsSubjectCacheInterface.ShutDownCacheManager"

End Sub

'--------------------------------------------------------------------------------
Public Property Get MaxArezzoAllowed() As Integer
'--------------------------------------------------------------------------------
'Retrieve the maximum number of Arezzo allowed to load by this server
'--------------------------------------------------------------------------------
Dim oACM As clsSubjectCacheManager
    
    Set oACM = New clsSubjectCacheManager
    
    MaxArezzoAllowed = oACM.MaxArezzoAllowed
    
    Set oACM = Nothing

End Property

'--------------------------------------------------------------------------------
Public Property Let MaxArezzoAllowed(ByVal nMaxArezzoAllowed As Integer)
'--------------------------------------------------------------------------------
' Set the maximum number of Arezzos allowed to load by the Cache Manager
' If nMaxArezzoAllowed < current MaxArezzoAllowed
' then the SCM will try to shut down "Not Busy" Arezzos if it can
'--------------------------------------------------------------------------------
Dim sglStartTime
Dim oACM As clsSubjectCacheManager
    
    On Error GoTo ErrLabel
    
    sglStartTime = Timer
    
    ' Keep trying
    Do While Timer < sglStartTime + msglPAUSE_TIME
        
        Set oACM = New clsSubjectCacheManager
        
        oACM.MaxArezzoAllowed = nMaxArezzoAllowed

        ' See if we succeeded
        ' Compare our target value with what the ACM actually managed to do
        If oACM.MaxArezzoAllowed = nMaxArezzoAllowed Then
            Exit Do
        End If
        
        ' We'll wait a bit and try again
        Set oACM = Nothing
    
        Sleep mlSLEEP_TIME
        
    Loop
    
    Set oACM = Nothing

Exit Property
    
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" & "clsSubjectCacheInterface.MaxArezzoAllowed( " & nMaxArezzoAllowed & ")"

End Property
''--------------------------------------------------------------------------------
'Public Sub CancelChanges(ByVal lStudyId As Long, _
'                            ByVal sSite As String, _
'                            ByVal lSubjectId As Long, _
'                            ByVal sDatabaseCnn As String)
''--------------------------------------------------------------------------------
'' Cancel changes to this subject, i.e. mark is as invalid
''--------------------------------------------------------------------------------
'Dim oACM As clsSubjectCacheManager
'
'    Set oACM = New clsSubjectCacheManager
'
'    Call oACM.MarkArezzoInvalid(lStudyId, sSite, lSubjectId, sDatabaseCnn)
'
'    Set oACM = Nothing
'
'End Sub

'------------------------------------------------------------------------------'
Public Sub ClearArezzoTokenTable(ByVal sConn As String)
'------------------------------------------------------------------------------'
'Deletes all entries in ArezzoToken table
'------------------------------------------------------------------------------'
Dim oACM As clsSubjectCacheManager
    
    Set oACM = New clsSubjectCacheManager
    
    Call oACM.ClearArezzoTokenTable(sConn)
    
    Set oACM = Nothing
    
End Sub

'------------------------------------------------------------------------------'
Public Function GetCacheReport() As Variant
'------------------------------------------------------------------------------'
' Returns an array giving information about each existing Cache Entry
'------------------------------------------------------------------------------'
Dim oACM As clsSubjectCacheManager

    Set oACM = New clsSubjectCacheManager

    GetCacheReport = oACM.GetCacheReport()

    Set oACM = Nothing
 
End Function

'------------------------------------------------------------------------------'
Public Function LoadEform(ByRef sSerialisedUser As String, ByVal lStudyId As Long, ByVal sSite As String, _
    ByVal lSubjectId As Long, ByVal sToken As String, ByVal lCRFPageTaskId As Long, ByRef sEFILockToken As String, _
    ByRef sVILockToken As String, ByRef bEFIUnavailable As Boolean, ByVal sDecimalPoint As String, _
    ByVal sThousandSeparator As String, Optional ByVal vAlerts As Variant, Optional ByVal vErrors As Variant, _
    Optional ByVal bAutoNext As Boolean = False) As String
'------------------------------------------------------------------------------'
' ic 18/08/2003
' function requests eform html from cache manager. called if the caller doesnt
' want to use the subject object directly due to com performance anomalies on the machine
' revisions
' ic 20/12/2004 bug 2395 - pass the sSerialisedUser byref so that lastused eform gets set and passed back
'------------------------------------------------------------------------------'
Dim oACM As clsSubjectCacheManager
    
    On Error GoTo handler
    
    Set oACM = New clsSubjectCacheManager
    LoadEform = oACM.LoadEformA(sSerialisedUser, lStudyId, sSite, lSubjectId, sToken, lCRFPageTaskId, sEFILockToken, _
        sVILockToken, bEFIUnavailable, sDecimalPoint, sThousandSeparator, vAlerts, vErrors, bAutoNext)
    Set oACM = Nothing
    Exit Function
    
handler:
    Err.Raise Err.Number, , Err.Description & "|" & "clsSubjectCacheInterface.LoadEform( " & lStudyId & "," & sSite & "," & lSubjectId & " )"
End Function

'------------------------------------------------------------------------------'
Public Function SaveEform(ByVal sSerialisedUser As String, ByVal lStudyId As Long, ByVal sSiteCode As String, _
    ByVal lSubjectId As Long, ByVal sToken As String, ByVal sCRFPageTaskId As String, ByVal sForm As String, _
    ByRef sEFILockToken As String, ByRef sVILockToken As String, ByVal bVReadOnly As Boolean, _
    ByVal bEReadOnly As Boolean, ByVal sLabCode As String, ByVal sDecimalPoint As String, _
    ByVal sThousandSeparator As String, ByRef sRegister As String, ByVal sLocalDate As String, _
    Optional ByVal nTimezoneOffset As Integer = 0) As Variant
'------------------------------------------------------------------------------'
' ic 18/08/2003
' function saves eform through cache manager. called if the caller doesnt
' want to use the subject object directly due to com performance anomalies on the machine
'------------------------------------------------------------------------------'
Dim oACM As clsSubjectCacheManager
    
    On Error GoTo handler
    
    Set oACM = New clsSubjectCacheManager
    SaveEform = oACM.SaveEformA(sSerialisedUser, lStudyId, sSiteCode, lSubjectId, sToken, sCRFPageTaskId, sForm, _
        sEFILockToken, sVILockToken, bVReadOnly, bEReadOnly, sLabCode, sDecimalPoint, sThousandSeparator, _
        sRegister, sLocalDate, nTimezoneOffset)
    Set oACM = Nothing
    Exit Function
 
handler:
   Err.Raise Err.Number, , Err.Description & "|" & "clsSubjectCacheInterface.SaveEform( " & lStudyId & "," & sSiteCode & "," & lSubjectId & " )"
End Function

'ic 16/03/2004 remove conditional compilation
'#If ORAMA = 1 Then
'------------------------------------------------------------------------------'
Public Function CheckArezzoEvents(ByVal sDatabaseCnn As String, ByVal lStudyId As Long, ByVal sSiteCode As String, ByVal lSubjectId As Long, _
    ByVal sToken As String, ByVal sDatabase As String, ByVal sEformPageTaskId As String, ByVal sNext As String, _
    ByRef bArezzoEvents As Boolean) As String
'------------------------------------------------------------------------------'
' ic 12/09/2003
' function checks arezzo events through cache manager. called if the caller doesnt
' want to use the subject object directly due to com performance anomalies on the machine
' ic 18/05/2005 issue 2560, pass database connection for locating cached subjects
'------------------------------------------------------------------------------'
    Dim oACM As clsSubjectCacheManager
    
    On Error GoTo handler
    
    Set oACM = New clsSubjectCacheManager
    CheckArezzoEvents = oACM.CheckArezzoEventsA(sDatabaseCnn, lStudyId, sSiteCode, lSubjectId, sToken, sDatabase, sEformPageTaskId, _
        sNext, bArezzoEvents)
    Set oACM = Nothing
    Exit Function
 
handler:
   Err.Raise Err.Number, , Err.Description & "|" & "clsSubjectCacheInterface.CheckArezzoEvents( " & lStudyId & "," & sSiteCode & "," & lSubjectId & " )"
End Function

'------------------------------------------------------------------------------'
Public Sub SaveArezzoEvents(ByVal sDatabaseCnn As String, ByVal lStudyId As Long, ByVal sSiteCode As String, ByVal lSubjectId As Long, _
    ByVal sToken As String, ByVal sForm As String)
'------------------------------------------------------------------------------'
'   12/09/2003
'   function wrapper for SaveArezzoEventsB
'   ic 18/05/2005 issue 2560, pass database connection for locating cached subjects
'------------------------------------------------------------------------------'
    Dim oACM As clsSubjectCacheManager
    
    On Error GoTo handler
    
    Set oACM = New clsSubjectCacheManager
    Call oACM.SaveArezzoEventsA(sDatabaseCnn, lStudyId, sSiteCode, lSubjectId, sToken, sForm)
    Set oACM = Nothing
    Exit Sub
 
handler:
   Err.Raise Err.Number, , Err.Description & "|" & "clsSubjectCacheInterface.SaveArezzoEvents( " & lStudyId & "," & sSiteCode & "," & lSubjectId & " )"
End Sub
'#End If
