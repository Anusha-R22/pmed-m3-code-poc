Attribute VB_Name = "modPSSSupport"
'------------------------------------------------------------------------------'
' File:     modPSSSupport.bas
' Copyright:    InferMed Ltd. 1999. All Rights Reserved
' Author:   Nicky Johns, September 1999
' Purpose:  Contains the routines for accessing the PSS
'           (Protocols Storage Services)
'------------------------------------------------------------------------------'
' Revisions:
'   NCJ 17 Sept 99 - Moved code here from ProformaEditor.bas
'   SDM 10/11/99    Copied in error routines
'   Mo Morris   17/1/00, changes made to StartUpPSS
'   Mo Morris   16/2/00     CheckProteditForTrial removed.
'   NCJ 10 Oct 00 - Ignore errors in DeleteProformaTrial
'   TA 09/10/2001: Changes to oracle opening parameters to fix bug in 2.2
'------------------------------------------------------------------------------'

Option Explicit

Public gpssProtocolStorage As ProtocolStorage
Public gpssProtocol As Protocol

'---------------------------------------------------------------------
Public Sub StartUpPSS()
'---------------------------------------------------------------------
'Make initial call to PSS.DLL
'and inform it of Macro's current database location
'Mo Morris 26/11/99
'OpenDatabase call changed after PSS.DLL converted to ADO (Access and SQL Server)
'Mo Morris 17/1/00
'the following security paramaters DatabaseUser, DatabasePassword, MacroUserName
'and MacroPassword are now additionally passed to PSS.DLL
'---------------------------------------------------------------------
Dim sCallParameters As String

    On Error GoTo ErrHandler
    
    
    'for oracle we have to pass the databasename as the server name as the PSS
    '   uses the server name as the TNS name when dealing with oracle.
     'ASH 12/9/2002 Changed gUser.DatabaseName to gUser.NameOfDatabase
    If gUser.DatabaseType = MACRODatabaseType.Oracle80 Then
        sCallParameters = gUser.DatabaseType & "|" & gUser.DatabasePath _
            & "|" & gUser.DatabaseName & "|" & gUser.NameOfDatabase _
            & "|" & gUser.DatabaseUser & "|" & gUser.DatabasePassword _
            & "|" & gUser.UserName & "|" & gUser.Password
    Else
        'ASH 12/9/2002 Changed gUser.DatabaseName to gUser.NameOfDatabase
        sCallParameters = gUser.DatabaseType & "|" & gUser.DatabasePath _
            & "|" & gUser.ServerName & "|" & gUser.NameOfDatabase _
            & "|" & gUser.DatabaseUser & "|" & gUser.DatabasePassword _
            & "|" & gUser.UserName & "|" & gUser.Password
    End If

    Set gpssProtocolStorage = New ProtocolStorage
    gpssProtocolStorage.OpenDatabase sCallParameters
    gpssProtocolStorage.Protocols.Build     ' Build the Protocols collection
    'Debug.Print "StartUpPSS after PSS.Build count=" & gpssProtocolStorage.Protocols.Count
   
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "StartUpPSS", "modPSSSupport")
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
Public Sub ShutDownPSS()
'---------------------------------------------------------------------
' Close down the PSS
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    Set gpssProtocol = Nothing
    Set gpssProtocolStorage = Nothing
   
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "ShutDownPSS", "modPSSSuport.bas")
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
Public Sub DeleteProformaTrial(ByVal sClinicalTrialName As String)
'---------------------------------------------------------------------
' Arezzo version is to delete from PSS database - NCJ 11/8/99
'---------------------------------------------------------------------

    ' NCJ 10/10/00 - Ignore errors in deleting an Arezzo protocol definition
    ' because it may genuinely not be there, e.g. if a GGB import is done twice
    ' without first creating the Arezzo file
    On Error Resume Next
    'Debug.Print "DeleteProformaTrial PSS Destroy call on " & sClinicalTrialName
    gpssProtocolStorage.Protocols.Destroy sClinicalTrialName
    
    On Error GoTo ErrHandler
    gpssProtocolStorage.Protocols.Build     ' Rebuild the Protocols collection
    'Debug.Print "DeleteProformaTrial after PSS.Build count=" & gpssProtocolStorage.Protocols.Count
   
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "DeleteProformaTrial(" & sClinicalTrialName & ")", _
                                "modPSSSupport")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
End Sub


