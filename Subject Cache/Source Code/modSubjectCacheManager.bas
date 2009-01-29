Attribute VB_Name = "modSubjectCacheManager"
'----------------------------------------------------------------------------
'   File:       modSubjectCacheManager.bas
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   Author:     Nicky Johns, July 2002 (based on original by Zulfi, Nov '01)
'   Purpose:    Responsible for creating and providing subject, study, responses etc
'-----------------------------------------------------------------------------
' Revisions:
'   NCJ 18-22 July - Initial development
'   NCJ 23 July 02 - Some tidying up following TA code review
'   MACRO 3.0
'   NCJ 17 Sept 02 - Updated for new MACRO 3.0 locking model
'   NCJ 24 Jan 03 - Added sCountry parameter to NewSubject
'----------------------------------------------------------------------------
Option Explicit

' This value MUST be stored in PUBLIC Variable HERE
' so that it is available to ALL users of this single-threaded DLL
Public gnMaxArezzoAllowed As Integer

Private Const msDELIMITER = "|"
' Values for columns in ArezzoToken database table
Private Const msAREZZO_COLS = "ArezzoID|DBToken|ClinicalTrialId|TrialSite|PersonId"
Private Const msAREZZO_TOKEN_COLS = "ArezzoID|ClinicalTrialId|TrialSite|PersonId"
Private Const msAREZZO_TOKEN = "ArezzoToken"
Private Const msCOL_AREZZOID = "ArezzoID"
Private Const msCOL_DBTOKEN = "DBToken"
Private Const msCOL_CLINICALTRIALID = "ClinicalTrialId"
Private Const msCOL_TRIALSITE = "TrialSite"
Private Const msCOL_PERSONID = "PersonId"

Private Const msCACHE_BUSY = "MACRO Subject Cache is busy"

''-----------------------------------------------------------------------------
'Public Sub MarkArezzoInvalidA(ByVal sToken As String, ByVal sConnString As String)
''-----------------------------------------------------------------------------
'' Call this function when user wants to cancel the changes made to subject
'' Change the ArezzoToken value in DB, so it will become invalid
''-----------------------------------------------------------------------------
'Dim oQD As QueryDef
'Dim oQS As QueryServer
'Dim vValue As Variant
'Dim sNewKey As String
'Dim sPrevKey As String
'
'    On Error GoTo ErrLabel
'
'    Set oQD = New QueryDef
'    Set oQS = New QueryServer
'
'    oQS.Init sConnString
'    oQS.ConnectionOpen
'
'    ' Save value of 0 for DBToken
'    oQD.InitSave msAREZZO_TOKEN, msCOL_DBTOKEN, 0, msCOL_AREZZOID, sToken
'
'    oQS.QueryUpdate oQD
'
'    oQS.ConnectionClose
'
'    Set oQS = Nothing
'    Set oQD = Nothing
'
'    Exit Sub
'
'ErrLabel:
'    Err.Raise Err.Number, , Err.Description & "|" & "modSubjectCacheManager.MarkArezzoInvalidA"
'End Sub

'-----------------------------------------------------------------------------
Public Function IsArezzoInvalid(ByVal sCacheToken As String, _
                                ByVal lStudyId As Long, _
                                ByVal sSite As String, _
                                ByVal lSubjectId As Long, _
                                ByVal sConnString As String) As Boolean
'-----------------------------------------------------------------------------
' Checks if Arezzo is invalid for this subject by querying through DB
' sCacheToken is the Arezzo token
' If the database does not contain our row, it means our Arezzo isn't valid
'-----------------------------------------------------------------------------
Dim oQD As QueryDef
Dim oQS As QueryServer
Dim vValue As Variant

    On Error GoTo ErrLabel

    Set oQD = New QueryDef
    Set oQS = New QueryServer
      
    oQS.Init sConnString
    oQS.ConnectionOpen

    ' Look for a matching row with our ArezzoToken
    oQD.InitSelect msAREZZO_TOKEN, Split(msAREZZO_TOKEN_COLS, msDELIMITER), _
                                   Split(msAREZZO_TOKEN_COLS, msDELIMITER), _
                                    Array(sCacheToken, lStudyId, sSite, lSubjectId)

    vValue = oQS.SelectArray(oQD)

    oQS.ConnectionClose

    Set oQS = Nothing
    Set oQD = Nothing

    ' If we found it, then it's still valid ('cos no-one else has blitzed it)
    If IsNull(vValue) Then
        ' No matching row found
        IsArezzoInvalid = True
    Else
        ' Our database row was found intact so we haven't been invalidated
        IsArezzoInvalid = False
    End If
      
Exit Function
    
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" _
            & "modSubjectCacheManager.IsArezzoInvalid( " & sCacheToken & " )"
    
End Function

''-----------------------------------------------------------------------------
'Public Function SaveArezzoToken(ByVal sArezzoID As String, _
'                           ByVal sConnString As String, _
'                           ByVal lStudyId As Long, _
'                           ByVal sSite As String, _
'                           Optional ByVal lSubjectId As Long, _
'                           Optional ByVal sPrevArezzoID = "") As Long
''-----------------------------------------------------------------------------
'' Saves the Arezzo information in the ArezzoToken database table
'' and returns the new DBToken
''-----------------------------------------------------------------------------
'Dim oQD As QueryDef
'Dim oQS As QueryServer
'Dim lLastToken As Long
'Dim vValues As Variant
'
'    On Error GoTo ErrLabel
'
'    Set oQD = New QueryDef
'    Set oQS = New QueryServer
'
'    If sPrevArezzoID = "" Then
'        CheckAndDelete sArezzoID, sConnString
'    Else
'        CheckAndDelete sPrevArezzoID, sConnString
'    End If
'
'    oQS.ConnectionOpen sConnString
'
'    'first try to get the largest DBToken value
'    oQD.InitNewId msAREZZO_TOKEN, msCOL_DBTOKEN
'
'    lLastToken = oQS.QueryDefNewId(oQD)
'
'    'try generating a test error
'    'Err.Raise "999", , "Forced error"
'
'    vValues = Array(sArezzoID, lLastToken, lStudyId, sSite, lSubjectId)
'
'    oQS.BeginTrans
'
'    oQD.InitSave msAREZZO_TOKEN, Split(msAREZZO_COLS, msDELIMITER), vValues
'
'    oQS.QueryInsert oQD
'
'    oQS.Commit
'
'    SaveArezzoToken = lLastToken
'
'    Set oQD = Nothing
'    Set oQS = Nothing
'
'    'debugging calls
''    WriteLog "[]DB Token = " & lLastToken
'
'    Exit Function
'
'ErrLabel:
'
'    Err.Raise Err.Number, , Err.Description & "|" & "modSubjectCacheManager.SaveArezzoToken"
'End Function
'
''-----------------------------------------------------------------------------
'Private Function CheckAndDelete(ByVal sArezzoID As String, ByVal sConnString As String)
''-----------------------------------------------------------------------------
'' Deletes any rows in ArezzoToken table for sArezzoID
'' Call this before adding a new one
''-----------------------------------------------------------------------------
'Dim oQD As QueryDef
'Dim oQS As QueryServer
'
'    On Error GoTo ErrLabel
'
'    Set oQD = New QueryDef
'    Set oQS = New QueryServer
'
'    oQD.InitDelete msAREZZO_TOKEN, msCOL_AREZZOID, sArezzoID
'
'    oQS.ConnectionOpen sConnString
'
'    oQS.BeginTrans
'
'    oQS.SelectDelete oQD
'
'    oQS.Commit
'
'    Set oQS = Nothing
'    Set oQD = Nothing
'
'    Exit Function
'
'ErrLabel:
'    Err.Raise Err.Number, , Err.Description & "|" & "modSubjectCacheManager.CheckAndDelete"
'
'End Function

''-------------------------------------------------------------------------
'Public Function ClearArezzoTokenTableA(ByVal sConnString As String)
''-------------------------------------------------------------------------
''clear ArezzoToken table if user has decided to shut down the server
''-------------------------------------------------------------------------
'Dim oQD As QueryDef
'Dim oQS As QueryServer
'Dim vArezzoIDs As Variant
'Dim i As Integer
'
'    On Error GoTo ErrLabel
'
'    Set oQD = New QueryDef
'    Set oQS = New QueryServer
'
'
'    oQD.InitSelect msAREZZO_TOKEN, msCOL_AREZZOID
'
'    oQS.ConnectionOpen sConnString
'
'    vArezzoIDs = oQS.SelectArray(oQD)
'
'    'delete all entries from the ArezzoToken table
'    If Not IsNull(vArezzoIDs) Then
'
'        For i = 0 To UBound(vArezzoIDs, 2)
'            CheckAndDelete vArezzoIDs(0, i), sConnString
'        Next i
'    End If
'
'    Set oQS = Nothing
'    Set oQD = Nothing
'
'    Exit Function
'
'ErrLabel:
'    Err.Raise Err.Number, , Err.Description & "|" & "modSubjectCacheManager.ClearArezzoTokenTableA"
'End Function

'--------------------------------------------------------------------------------
Public Function LoadSubjectA(ByVal sDatabaseCnn As String, _
                        ByVal lStudyId As Long, _
                        ByVal sSite As String, _
                        ByVal lSubjectId As Long, _
                        ByVal sUserName As String, _
                        ByVal nUpdateMode As Integer, _
                        ByVal sUserNameFull As String, _
                        ByVal sUserRole As String) As Variant
'--------------------------------------------------------------------------------
' Load specified subject and return array containing:
'   Result code (0 if OK, 1 if no Cache entries available, 2 if Study is locked)
'   New StudySubject object (if Result code = 0)
'   CacheToken for the subject if Result code = 0, OR the text result message
' NB The CacheToken must be used in ReleaseSubject
'--------------------------------------------------------------------------------
Dim arrSubject() As Variant
Dim oSubject As StudySubject

    On Error GoTo ErrLabel
    
    ReDim arrSubject(2)
    
    ' Set "Not OK" to begin with
    arrSubject(0) = eCacheLoadResult.clrNoCacheEntries
    
    ' Call Nicky's new LoadSubjectB
    Set oSubject = LoadSubjectB(sDatabaseCnn, lStudyId, sSite, lSubjectId, sUserName, nUpdateMode, _
                            sUserNameFull, sUserRole)

    If Not oSubject Is Nothing Then
        ' Deal with study being locked
        If oSubject.CouldNotLoad Then
            arrSubject(0) = eCacheLoadResult.clrStudyLocked
            arrSubject(2) = oSubject.CouldNotLoadReason
        Else
            ' Set the "OK" flag
            arrSubject(0) = eCacheLoadResult.clrOK
            Set arrSubject(1) = oSubject
            arrSubject(2) = oSubject.CacheToken
        End If
    Else
        arrSubject(2) = msCACHE_BUSY
    End If
    
    Set oSubject = Nothing
    
    LoadSubjectA = arrSubject
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" _
                    & "modSubjectCacheManager.LoadSubjectA"

End Function

'--------------------------------------------------------------------------------------------------
Public Function CreateNewSubjectA(ByVal sUserName As String, _
                                 ByVal sDatabaseCnn As String, _
                                 ByVal lStudyId As Long, _
                                 ByVal sSite As String, _
                                 ByVal sCountry As String, _
                                ByVal sUserNameFull As String, _
                                ByVal sUserRole As String) As Variant
'--------------------------------------------------------------------------------------------------
' Create new subject and return array containing:
'   Result code (0 if OK, 1 if no cache entries available, 2 if Study is locked)
'   New StudySubject object (if Result code = 0)
'   CacheToken for the subject if Result code = 0, OR the text result message
' NB The CacheToken must be used in ReleaseSubject
' NCJ 24 Jan 03 - Added sCountry parameter
' NCJ 6 May 03 - Added new params. sUserNameFull and sUserRole
'--------------------------------------------------------------------------------------------------
Dim arrSubject() As Variant
Dim oSubject As StudySubject

    On Error GoTo ErrLabel
    
    ReDim arrSubject(2)
    
    ' Set "Not OK" to begin with
    arrSubject(0) = eCacheLoadResult.clrNoCacheEntries
    
    Set oSubject = CreateNewSubjectB(sUserName, sDatabaseCnn, lStudyId, sSite, sCountry, _
                            sUserNameFull, sUserRole)

    If Not oSubject Is Nothing Then
        ' Deal with study being locked
        If oSubject.CouldNotLoad Then
            arrSubject(0) = eCacheLoadResult.clrStudyLocked
            arrSubject(2) = oSubject.CouldNotLoadReason
        Else
            ' Set the "OK" flag
            arrSubject(0) = eCacheLoadResult.clrOK
            Set arrSubject(1) = oSubject
            arrSubject(2) = oSubject.CacheToken
        End If
    Else
        arrSubject(2) = msCACHE_BUSY
    End If
    
    Set oSubject = Nothing
    
    CreateNewSubjectA = arrSubject
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|" _
                    & "modSubjectCacheManager.CreateNewSubjectA"

End Function

