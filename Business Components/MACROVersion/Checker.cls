VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Checker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------------------'
'   File:       Checker.cls
'   Copyright:  InferMed Ltd. 2005-2007. All Rights Reserved
'   Author:     Ian Curtis, October 2005
'   Purpose:    Macro version checking class
'----------------------------------------------------------------------------------------'
' Revisions:
' ic 02/10/2006  issue 2813 the about dialog should list the patch version
' ic 23/02/2007 bug 2855 now that all databases have been updated with clinical coding
'               changes, we check to see if the clinical coding dlls are installed to
'               determine whether clinical coding is activated instead of checking for
'               the existance of a certain table
' NCJ 9 Oct 07 - Use Settings file for IsMedDRA (Patch 3.0.78CC)
'----------------------------------------------------------------------------------------'

Option Explicit
Private moConnection As Connection
Private Const mCCSwitch As String = "CLINICALCODING"
Private Const mActiveDirectorySwitch As String = "ACTIVEDIRECTORY"

' ic 02/10/2006  issue 2813 the about dialog should list the patch version
'change this constant value when patch version changes
Private Const mCurrentPatchVersion As String = "3.0.82"

'----------------------------------------------------------------------------------------'
Public Function PatchVersion() As String
'----------------------------------------------------------------------------------------'
' ic 02/10/2006
' issue 2813 the about dialog should list the patch version
' return the current patch version
'----------------------------------------------------------------------------------------'
    PatchVersion = mCurrentPatchVersion
End Function


'----------------------------------------------------------------------------------------'
Public Function HasUpgrade(sCon As String, sUpgrade As String)
'----------------------------------------------------------------------------------------'
' ic
' check whether a database has an upgrade
' revisions
' ic 23/02/2007 bug 2855 dont pass the connection string to the IsMedDRA function
'----------------------------------------------------------------------------------------'
Dim b As Boolean
    
    b = False
    
    Select Case sUpgrade
    Case mCCSwitch:
        b = IsMedDRA()
    Case mActiveDirectorySwitch:
        b = IsActiveDirectory(sCon)
    Case Else:
    End Select
    
    HasUpgrade = b
End Function

'----------------------------------------------------------------------------------------'
Private Function IsMedDRA()
'----------------------------------------------------------------------------------------'
' ic
' does database have meddra coding upgrade. requires a macro db connection string
' (not security db)
' revisions
' ic 23/02/2007 bug 2855 now that all databases have been updated with clinical coding
'               changes, we check to see if the clinical coding dlls are installed to
'               determine whether clinical coding is activated instead of checking for
'               the existance of a certain table
' NCJ 9 Oct 07 - Use settings file instead
'----------------------------------------------------------------------------------------'
Dim bInstalled As Boolean

    ' bUsingWWWDll = TRUE because we're a DLL!
    Call InitialiseSettingsFile(True)
    IsMedDRA = (LCase(GetMACROSetting(MACRO_CLINICAL_CODING, "false")) = "true")

'    bInstalled = False
'    On Error GoTo NotInstalled
'
'    'try to create a clinical coding business services dll, if it fails cc is not installed
'    Dim oCCBS As MACROCCBS30.CodedTerm
'    Set oCCBS = New MACROCCBS30.CodedTerm
'    Set oCCBS = Nothing
'    bInstalled = True
'
'NotInstalled:
'    IsMedDRA = bInstalled
'

End Function

'----------------------------------------------------------------------------------------'
Private Function IsActiveDirectory(sSecurityCon As String)
'----------------------------------------------------------------------------------------'
' ic
' does database have active directory coding upgrade. requires a macro security db
' connection string
'----------------------------------------------------------------------------------------'
Dim rs As Recordset
Dim sSQL As String
Dim bIsActiveDirectory As Boolean


    On Error GoTo Ignore
    bIsActiveDirectory = False
    
    'open the connection to the macro database
    ConnectionOpen (sSecurityCon)
    
    'create sql to see if theres a ACTIVEDIRECTORYSERVERS table
    sSQL = "SELECT TABLE_NAME FROM "
    If (InStr(sSecurityCon, "SQLOLEDB") > 0) Then
        'sql server database
        sSQL = sSQL & "INFORMATION_SCHEMA.Tables "
    Else
        'oracle database
        sSQL = sSQL & "USER_TABLES "
    End If
    sSQL = sSQL & "WHERE TABLE_NAME = 'ACTIVEDIRECTORYSERVERS'"
    
    'get the data
    Set rs = New Recordset
    With rs
        .Open sSQL, moConnection, adOpenKeyset, adLockPessimistic
        If (Not (.BOF And .EOF)) Then
            bIsActiveDirectory = True
        End If
    End With
    
Ignore:
    Call ConnectionClose
    IsActiveDirectory = bIsActiveDirectory
End Function
'----------------------------------------------------------------------
Private Sub ConnectionOpen(sCon As String)
'----------------------------------------------------------------------
' Open connection.
'----------------------------------------------------------------------
    If (moConnection Is Nothing) Then
        Set moConnection = New Connection
        moConnection.Open sCon
        moConnection.CursorLocation = adUseClient
    End If
End Sub

'----------------------------------------------------------------------------------------'
Private Sub ConnectionClose()
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'
    If (Not (moConnection Is Nothing)) Then
        moConnection.Close
        Set moConnection = Nothing
    End If
End Sub


