Attribute VB_Name = "modCacheTokens"
'----------------------------------------------------------------------------
'   File:       modCacheTokens
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   Author:     Nicky Johns, July 2002
'   Purpose:    Handling of Arezzo tokens for the Cache Manager
'-----------------------------------------------------------------------------
' Revisions:
'   NCJ 18-22 July - Initial development
'   NCJ 19 Sept 02 - This version for MACRO 3.0
'----------------------------------------------------------------------------

Private Const msSEPARATOR = ","

Public Enum eCacheLoadResult
    clrOK = 0
    clrNoCacheEntries = 1
    clrStudyLocked = 2
End Enum

Option Explicit

'--------------------------------------------------------------------------------------------------
Public Sub UnwrapToken(ByVal sToken As String, _
                            ByRef lStudyId As Long, _
                            ByRef sSite As String, _
                            ByRef lSubjectId As Long)
'--------------------------------------------------------------------------------------------------
' Split token into its component parts
' Returns its three values
'--------------------------------------------------------------------------------------------------
Dim vTok As Variant

    vTok = Split(sToken, msSEPARATOR)
    lStudyId = CLng(vTok(0))
    sSite = vTok(1)
    lSubjectId = CLng(vTok(2))
    
End Sub

'--------------------------------------------------------------------------------------------------
Public Function GetToken(ByVal lStudyId As Long, _
                            ByVal sSite As String, _
                            ByVal lSubjectId As Long) As String
'--------------------------------------------------------------------------------------------------
' Create a token from three given values
'--------------------------------------------------------------------------------------------------

    GetToken = lStudyId & msSEPARATOR & sSite & msSEPARATOR & lSubjectId
    
End Function

'------------------------------------------------------------------------------'
Public Sub WriteLog(sLog As String)
'------------------------------------------------------------------------------'
' Write errors to the log file. Assume that "SCMLog.dat" already exist in temp
' folder under the application (if not, no logging happens)
'------------------------------------------------------------------------------'
Dim objFile As TextStream
Dim objFSO As Scripting.FileSystemObject
Dim sFileName As String

    On Error GoTo IgnoreErrors
    
    Set objFSO = New Scripting.FileSystemObject
    
    sFileName = App.Path & "\temp\SCMLog.dat"
    If objFSO.FileExists(sFileName) Then
        Set objFile = objFSO.OpenTextFile(sFileName, ForAppending)
        objFile.WriteLine Format(Now, "hh:mm:ss") & " * " & sLog
        objFile.Close
    End If
    
IgnoreErrors:
    Set objFSO = Nothing
        
End Sub

