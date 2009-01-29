VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UpSecurity20To22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2003. All Rights Reserved
'   File:       UpSecurity20To22.cls
'   Author:     Richard Meinesz 18/06/2003
'   Purpose:    Class to upgrade a MACRO secuirty database to the latest version of MACRO 2.2 from any previous version
'----------------------------------------------------------------------------------------'
'REVISIONS:
'
'----------------------------------------------------------------------------------------'

Option Explicit

Private moSecurityCon As ADODB.Connection
Private meDatabaseType As MACRODatabaseType

'----------------------------------------------------------------------------------------'
Public Sub Init(oSecCon As ADODB.Connection, eDatabasrType As MACRODatabaseType)
'----------------------------------------------------------------------------------------'
'REM 18/06/03
'Initialise the secuirty connection
'----------------------------------------------------------------------------------------'
    
    Set moSecurityCon = oSecCon
    meDatabaseType = eDatabasrType
End Sub

'---------------------------------------------------------------------
Public Sub UpgradeSecurityDatabaseTo22()
'---------------------------------------------------------------------
' Check for necessary upgrading of the Security.mdb
'---------------------------------------------------------------------

    ' Upgrade 2.0 Security database
    Call UpgradeSecurityDatabase2_0
    ' Upgrade 2.1 Security database
    Call UpGradeSecurityDatabase2_1
    ' Upgrade 2.2 Security database
    Call UpGradeSecurityDatabase2_2
    
End Sub

'---------------------------------------------------------------------
Private Sub UpgradeSecurityDatabase2_0()
'---------------------------------------------------------------------
' Upgrade a 2.0 Security.mdb
'---------------------------------------------------------------------
Dim rsTemp As ADODB.Recordset
Dim sSQL As String
Dim sMacroVersion As String
Dim sBuildSubVersion As String

    On Error GoTo ErrHandler

    sSQL = "SELECT * FROM SecurityControl"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, moSecurityCon, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        MsgBox ("Your Macro Security database is not valid. Macro is being closed down.")
        ExitMACRO
        MACROEnd
        Exit Sub
    End If
    
    sMacroVersion = rsTemp![MACROVersion]
    If sMacroVersion <> "2.0" Then
        Exit Sub
    End If
    
    sBuildSubVersion = rsTemp![BuildSubVersion]
    
    'Check for need to upgrade from BuildSubVersion [15] to [16]
    If sBuildSubVersion = "15" Then
        UpGradeSecurity15to16
        sBuildSubVersion = "16"
    End If
    
    'Check for need to upgrade from BuildSubVersion [16] to [17]
    If sBuildSubVersion = "16" Then
        sBuildSubVersion = "17"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'Check for need to upgrade from BuildSubVersion [17] to [18]
    If sBuildSubVersion = "17" Then
        sBuildSubVersion = "18"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'Check for need to upgrade from BuildSubVersion [18] to [19]
    If sBuildSubVersion = "18" Then
        sBuildSubVersion = "19"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'Check for need to upgrade from BuildSubVersion [19] to [20]
    If sBuildSubVersion = "19" Then
        UpGradeSecurity19to20   ' Special upgrade
        sBuildSubVersion = "20"
    End If

    'Check for need to upgrade from BuildSubVersion [20] to [21]
    If sBuildSubVersion = "20" Then
        sBuildSubVersion = "21"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If

    'Check for need to upgrade from BuildSubVersion [21] to [22]
    ' NCJ 8/3/00
    If sBuildSubVersion = "21" Then
        sBuildSubVersion = "22"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'Check for need to upgrade from BuildSubVersion [22] to [23]
    ' WIllC 14/3/00
    If sBuildSubVersion = "22" Then
        sBuildSubVersion = "23"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'Check for need to upgrade from BuildSubVersion [23] to [27]
    ' NCJ 3/4/00
    If sBuildSubVersion = "23" Then
        sBuildSubVersion = "27"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If

    'Check for need to upgrade from BuildSubVersion [27] to [29]
    ' WIllC 26/4/00
    If sBuildSubVersion = "27" Then
        UpGradeSecurity27to29
        sBuildSubVersion = "29"
    End If
    
    If sBuildSubVersion = "29" Then
        sBuildSubVersion = "31"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If

    If sBuildSubVersion = "31" Then
        sBuildSubVersion = "32"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If

    'WillC 16/5/00
    If sBuildSubVersion = "32" Then
        UpGradeSecurity32to33
        sBuildSubVersion = "33"
    End If

    'WillC 26/5/00
    If sBuildSubVersion = "33" Then
        UpGradeSecurity33to34
        sBuildSubVersion = "34"
    End If

    'WillC 30/5/00
    If sBuildSubVersion = "34" Then
        UpGradeSecurity34to35
        sBuildSubVersion = "35"
    End If
    
    'NCJ 2/6/00
    If sBuildSubVersion = "35" Then
        sBuildSubVersion = "36"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
        'WillC 9/6/00
    If sBuildSubVersion = "36" Then
        UpGradeSecurity36to37
        sBuildSubVersion = "37"
    End If
    
        'Nicky 15/6/00
    If sBuildSubVersion = "37" Then
        sBuildSubVersion = "39"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If

        'Nicky 16/6/00
    If sBuildSubVersion = "39" Then
        sBuildSubVersion = "40"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpgradeSecurityDatabase", "modUpgradeDatabases.bas")
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
Private Sub UpGradeSecurity15to16()
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandler

    'check for need to remove a trailing space in the RoleFunction table
    sSQL = "SELECT * FROM RoleFunction" _
        & " WHERE RoleCode = 'MacroUser'" _
        & " AND FunctionCode = 'F3002 '"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, moSecurityCon, adOpenKeyset, adLockPessimistic, adCmdText
    If rsTemp.RecordCount = 1 Then
        rsTemp!FunctionCode = "F3002"
        rsTemp.Update
        'Debug.Print "F3002 changed"
    End If
    rsTemp.Close
    Set rsTemp = Nothing
    
    'Upgrade BuildSubVersion from [15] to [16]
    sSQL = "UPDATE SecurityControl Set BuildSubVersion = '16'"
    moSecurityCon.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpGradeSecurity15to16", "modUpgradeDatabases.bas")
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
Private Sub UpGradeSecurity19to20()
'---------------------------------------------------------------------
' NCJ 16/2/00 Ignore errors in this routine because we unfortunately
' got an incompatibility between upgraded 19 DBs and script-generated 19 DBs
' (this change should have been in UpGradeSecurity18to19 but we missed the boat)
'---------------------------------------------------------------------
Dim sSQL As String

    On Error Resume Next
    
    sSQL = "Insert INTO Function (FunctionCode,Function) " _
        & " VALUES ('F2002','Disable user')"
    moSecurityCon.Execute sSQL, dbFailOnError
    
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) " _
        & " VALUES ('MacroUser','F2002')"
    moSecurityCon.Execute sSQL, dbFailOnError

    On Error GoTo ErrHandler
    
    'Upgrade BuildSubVersion from [19] to [20]
    sSQL = "UPDATE SecurityControl Set BuildSubVersion = '20'"
    moSecurityCon.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpGradeSecurity19to20", "modUpgradeDatabases.bas")
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
Private Sub UpGradeSecurity27to29()
'---------------------------------------------------------------------
'WillC 26/4/00
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    'Change the Function descriptions replace the word trial with study
    sSQL = " UPDATE Function SET Function = 'Assign user to study' WHERE FunctionCode = 'F2006'"
    moSecurityCon.Execute sSQL, dbFailOnError
    sSQL = " UPDATE Function SET Function = 'Add site to study or study to site' WHERE FunctionCode = 'F4002'"
    moSecurityCon.Execute sSQL, dbFailOnError
    sSQL = " UPDATE Function SET Function = 'Remove site from study' WHERE FunctionCode = 'F4003'"
    moSecurityCon.Execute sSQL, dbFailOnError
    sSQL = " UPDATE Function SET Function = 'Change study status' WHERE FunctionCode = 'F4005'"
    moSecurityCon.Execute sSQL, dbFailOnError
    
    'Add 4 new functions to the function table
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F5013','View question audit trail')"
    moSecurityCon.Execute sSQL, dbFailOnError
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F5014','Overrule discrepancies')"
    moSecurityCon.Execute sSQL, dbFailOnError
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F5015','Add data entry question comment')"
    moSecurityCon.Execute sSQL, dbFailOnError
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F5016','View data entry question comments')"
    moSecurityCon.Execute sSQL, dbFailOnError
    
    
    ' Add the 4 new functions to the MacroUser role in the RoleFunction table.
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F5013')"
    moSecurityCon.Execute sSQL, dbFailOnError
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F5014')"
    moSecurityCon.Execute sSQL, dbFailOnError
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F5015')"
    moSecurityCon.Execute sSQL, dbFailOnError
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F5016')"
    moSecurityCon.Execute sSQL, dbFailOnError
    
           
    'Upgrade BuildSubVersion from [27] to [29]
    sSQL = "UPDATE SecurityControl Set BuildSubVersion = '29'"
    moSecurityCon.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpGradeSecurity27to29", "modUpgradeDatabases.bas")
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
Private Sub UpGradeSecurity32to33()
'---------------------------------------------------------------------
' WillC 16/5/00
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

    'Upgrade BuildSubVersion from [32] to [33]
    sSQL = "UPDATE SecurityControl Set BuildSubVersion = '33'"
    moSecurityCon.Execute sSQL

    'Update the function text
    sSQL = "UPDATE Function Set Function = 'Access Data Review' Where FunctionCode = 'F1006'"
    moSecurityCon.Execute sSQL
    
   ' WillC 18/5/00 New Functions
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F5017','Create Discrepancy')"
    moSecurityCon.Execute sSQL
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F5018','Create SDV Remark')"
    moSecurityCon.Execute sSQL
    
    
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F5017')"
    moSecurityCon.Execute sSQL
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F5018')"
    moSecurityCon.Execute sSQL

    

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeSecurity32to33", "modUpgradeDatabases.bas")
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
Private Sub UpGradeSecurity33to34()
'---------------------------------------------------------------------
' WillC 26/5/00
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsRoleFunction As ADODB.Recordset

    On Error GoTo ErrHandler
    
    'Upgrade BuildSubVersion from [33] to [34]
    sSQL = "UPDATE SecurityControl Set BuildSubVersion = '34'"
    moSecurityCon.Execute sSQL

    'Update the function text
    sSQL = "UPDATE Function Set Function = 'System integrity check' Where FunctionCode = 'F5010'"
    moSecurityCon.Execute sSQL
    
    sSQL = "UPDATE Function Set Function = 'Audit trail integrity check' Where FunctionCode = 'F5012'"
    moSecurityCon.Execute sSQL
   
   'Insert New Function
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F3023','Update Arezzo from GGB import')"
    moSecurityCon.Execute sSQL
        
        
    'Delete from RoleFunction
    sSQL = "DELETE * FROM RoleFunction WHERE FunctionCode = 'F5011'"
    moSecurityCon.Execute sSQL
    
    'Open the RoleFunction table and use the recordset.update method to clear the deletion
    'out of the table so we can do delete to the function table without error further  below.
    sSQL = "SELECT * From RoleFunction "
    Set rsRoleFunction = New ADODB.Recordset
    rsRoleFunction.Open sSQL, moSecurityCon, adOpenDynamic, adLockOptimistic, adCmdText
    With rsRoleFunction
        .MoveFirst
        .MoveLast
        .Update
     If .RecordCount Then
     End If
        .Close
    End With
    Set rsRoleFunction = Nothing
        
    'Delete from Function
    sSQL = "DELETE  FROM Function WHERE FunctionCode = 'F5011'"
    moSecurityCon.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeSecurity33to34", "modUpgradeDatabases")
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
Private Sub UpGradeSecurity34to35()
'---------------------------------------------------------------------
' WIllC 30 May 00
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

    sSQL = "UPDATE Function Set Function = 'Create SDV Mark' Where FunctionCode = 'F5018'"
    moSecurityCon.Execute sSQL

    'Upgrade BuildSubVersion from [34] to [35]
    sSQL = "UPDATE SecurityControl Set BuildSubVersion = '35'"
    moSecurityCon.Execute sSQL

    
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeSecurity34to35", "modUpgradeDatabases")
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
Private Sub UpGradeSecurity36to37()
'---------------------------------------------------------------------
' WIllC 9 June 00
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

   'Insert New Function
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F5019','Use WORD Templates')"
    moSecurityCon.Execute sSQL

    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F5019')"
    moSecurityCon.Execute sSQL

    'Upgrade BuildSubVersion from [36] to [37]
    sSQL = "UPDATE SecurityControl Set BuildSubVersion = '37'"
    moSecurityCon.Execute sSQL

    
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeSecurity36to37", "modUpgradeDatabases")
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
Private Sub UpGradeSecurityToSubVersion(sSubVersion As String)
'---------------------------------------------------------------------
' Nicky 1/9/00
' Generic routine to update the build subversion in SecurityControl in Security.mdb
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

    sSQL = "UPDATE SecurityControl Set BuildSubVersion = '" & sSubVersion & "'"
    moSecurityCon.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeSecurityToSubVersion" & sSubVersion, "modUpgradeDatabases.bas")
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
Private Sub UpGradeSecurityDatabase2_1()
'---------------------------------------------------------------------
' Upgrade a 2.1 Security database, checking first for upgrade from 2.0 to 2.1
'---------------------------------------------------------------------

Dim rsTemp As ADODB.Recordset
Dim sSQL As String
Dim sMacroVersion As String
Dim sBuildSubVersion As String
Dim sMsg As String

    On Error GoTo ErrHandler

    sSQL = "SELECT * FROM SecurityControl"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, moSecurityCon, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        MsgBox ("Your Macro Security database is not valid. Macro is being closed down.")
        ExitMACRO
        MACROEnd
        Exit Sub
    End If
    
    ' Pick up the database version
    sMacroVersion = rsTemp![MACROVersion]
    sBuildSubVersion = rsTemp![BuildSubVersion]
    rsTemp.Close
    Set rsTemp = Nothing
    
    'don't do anything if futre version
    If Val(sMacroVersion) >= Val(2.2) Then
        Exit Sub
    End If
    ' Check for version 2.0
    If sMacroVersion = "2.0" Then
        'This is a temporary measure and must be taken out on everybody changing to 2.1
        'sMsg = "You are about to upgrade your MACRO security database from 2.0 to 2.1. Do you wish to continue?"
        'Select Case MsgBox(sMsg, vbQuestion + vbYesNo, gsDIALOG_TITLE)
            'Case vbYes
                ' Do initial upgrade from 2.0.x to 2.1.1
                UpGradeSecurity2_0to2_1_1
                sMacroVersion = "2.1"
                sBuildSubVersion = "1"
            'Case vbNo
                'Call ExitMACRO
                'Call MACROEnd
        'End Select
    End If
   
    

    
    If sMacroVersion <> "2.1" Then
        MsgBox ("Your Macro Security database is not valid. Macro is being closed down.")
        ExitMACRO
        MACROEnd
    End If
    
    'ADD NEW 2.1 UPGRADES HERE
    'Check for "01" as well as "1" only necessary for Sub Version 1
    If sBuildSubVersion = "1" Or sBuildSubVersion = "01" Then
        UpGradeSecurity2_1from1to2
        sBuildSubVersion = "2"
    End If
   
    If sBuildSubVersion = "2" Then
        sBuildSubVersion = "3"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    If sBuildSubVersion = "3" Then
        sBuildSubVersion = "4"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    If sBuildSubVersion = "4" Or sBuildSubVersion = "5" Then
        ' We missed out 5
        sBuildSubVersion = "6"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
   
   ' NCJ 1/11/00 - Upgrade 6 to 11
   ' NCJ 28/11/00 - Include 7,8,9,10,11,12 as well!
    Select Case sBuildSubVersion
    Case "6", "7", "8", "9", "10", "11", "12"
        sBuildSubVersion = "13"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    Case Else
        ' Nothing to do here
    End Select
    
    'Mo Morris 29/11/00 - Upgrade 13 to 14
    If sBuildSubVersion = "13" Then
        UpGradeSecurity2_1from13to14
        sBuildSubVersion = "14"
    End If
    
    'Mo Morris 19/12/00 - Upgrade 14,15,16 to 17
    Select Case sBuildSubVersion
    Case "14", "15", "16"
        sBuildSubVersion = "17"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    Case Else
        ' Nothing to do here
    End Select
    
    'TA 30/1/2001: Upgrade from 17 to 30 ' leaving room for Ronald builds
    If Val(sBuildSubVersion) >= 17 And Val(sBuildSubVersion) < 30 Then
        sBuildSubVersion = "30"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    ' NCJ 8 Feb 01 Upgrade to 32
    ' NCJ 19 Feb 01 - Upgrade 30-33 to 34
    ' NCJ 8 Mar 01 - Upgrade 34 and 35 to 36
   ' NCJ 23/3/01 - Upgrade to 37
   ' NCJ 2/4/01 - Upgrade to 38
   ' NCJ 10/4/01 - Upgrade to 39
    If Val(sBuildSubVersion) >= 30 And Val(sBuildSubVersion) <= 38 Then
        sBuildSubVersion = "39"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'TA 18/04/2001 - Upgrade 39 to 40
    If sBuildSubVersion = "39" Then
        UpGradeSecurity2_1from39to40
        sBuildSubVersion = "40"
    End If


Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeSecurityDatabase2_1", "modUpgradeDatabases")
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
Private Sub UpGradeSecurity2_0to2_1_1()
'---------------------------------------------------------------------
' WillC 1/8/00 SR3728
' This upgrades a 2.0.x (latest) Security database to 2.1.1
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    'Upgrade MACROVersion from [2.0] to [2.1]
    sSQL = "UPDATE SecurityControl Set MACROVersion = '2.1'"
    moSecurityCon.Execute sSQL

    'Upgrade BuildSubVersion to 1
    sSQL = "UPDATE SecurityControl Set BuildSubVersion = '1'"
    moSecurityCon.Execute sSQL
    
    ' Add new permissions
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F2009','Change System Properties')"
    moSecurityCon.Execute (sSQL)
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F2010','View System Log')"
    moSecurityCon.Execute (sSQL)
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F2011','Reset Password')"
    moSecurityCon.Execute (sSQL)
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F5020','View Subject Data')"
    moSecurityCon.Execute (sSQL)
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F2012','View Site/Server Communication')"
    moSecurityCon.Execute (sSQL)
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F2013','Restore Database')"
    moSecurityCon.Execute (sSQL)

    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F2009')"
    moSecurityCon.Execute (sSQL)
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F2010')"
    moSecurityCon.Execute (sSQL)
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F2011')"
    moSecurityCon.Execute (sSQL)
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F5020')"
    moSecurityCon.Execute (sSQL)
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F2012')"
    moSecurityCon.Execute (sSQL)
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F2013')"
    moSecurityCon.Execute (sSQL)

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeSecurity2_0to2_1_1", "modUpgradeDatabases")
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
Private Sub UpGradeSecurity2_1from1to2()
'---------------------------------------------------------------------
Dim sSQL As String


    On Error GoTo ErrHandler
    
    ' Add new permissions
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F6001','Maintain Laboratories')"
    moSecurityCon.Execute (sSQL)
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F6002','Maintain CTC Schemes')"
    moSecurityCon.Execute (sSQL)
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F6003','Maintain Clinical Tests')"
    moSecurityCon.Execute (sSQL)
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F6004','Maintain Normal Ranges')"
    moSecurityCon.Execute (sSQL)
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F6005','Maintain Common Toxicity Criteria')"
    moSecurityCon.Execute (sSQL)

    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F6001')"
    moSecurityCon.Execute (sSQL)
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F6002')"
    moSecurityCon.Execute (sSQL)
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F6003')"
    moSecurityCon.Execute (sSQL)
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F6004')"
    moSecurityCon.Execute (sSQL)
    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F6005')"
    moSecurityCon.Execute (sSQL)
    
    'Upgrade BuildSubVersion to 2
    sSQL = "UPDATE SecurityControl Set BuildSubVersion = '2'"
    moSecurityCon.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeSecurity2_1from1to2", "modUpgradeDatabases")
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
Private Sub UpGradeSecurity2_1from13to14()
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    ' Add new permissions
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F3024','Maintain Registration')"
    moSecurityCon.Execute (sSQL)

    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F3024')"
    moSecurityCon.Execute (sSQL)
    
    'Upgrade BuildSubVersion to 14
    sSQL = "UPDATE SecurityControl Set BuildSubVersion = '14'"
    moSecurityCon.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeSecurity2_1from13to14", "modUpgradeDatabases")
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
Private Sub UpGradeSecurity2_1from39to40()
'---------------------------------------------------------------------
Dim sSQL As String


    On Error GoTo ErrHandler

    'TA 18/04/2001: new Access Create Data Views function
    sSQL = "INSERT INTO Function (FunctionCode,Function) VALUES ('F1007','Access Create Data Views')"
    moSecurityCon.Execute sSQL
    sSQL = "INSERT INTO Rolefunction (RoleCode,FunctionCode) VALUES ('MacroUser','F1007')"
    moSecurityCon.Execute sSQL
    
    'Upgrade BuildSubVersion to 40
    sSQL = "UPDATE SecurityControl Set BuildSubVersion = '40'"
    moSecurityCon.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpGradeSecurity2_1from39to40", "modUpgradeDatabases.bas")
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
Private Sub UpGradeSecurityDatabase2_2()
'---------------------------------------------------------------------
' Upgrade a 2.1 Security database, checking first for upgrade from 2.0 to 2.1
'---------------------------------------------------------------------

Dim rsTemp As ADODB.Recordset
Dim sSQL As String
Dim sMacroVersion As String
Dim sBuildSubVersion As String
Dim sMsg As String

    On Error GoTo ErrHandler

    sSQL = "SELECT * FROM SecurityControl"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, moSecurityCon, adOpenKeyset, adLockReadOnly, adCmdText
    
    If rsTemp.RecordCount = 0 Then
        MsgBox ("Your Macro Security database is not valid. Macro is being closed down.")
        ExitMACRO
        MACROEnd
        Exit Sub
    End If
    
    ' Pick up the database version
    sMacroVersion = rsTemp![MACROVersion]
    sBuildSubVersion = rsTemp![BuildSubVersion]
    rsTemp.Close
    Set rsTemp = Nothing
    
    'don't do anything if futre version
    If Val(sMacroVersion) > Val(2.2) Then
        Exit Sub
    End If
    
    ' Check for version 2.1
    If sMacroVersion = "2.1" Then
        'sMsg = "You are about to upgrade your MACRO security database from 2.1 to 2.2. Do you wish to continue?"
        'Select Case MsgBox(sMsg, vbQuestion + vbYesNo, gsDIALOG_TITLE)
            'Case vbYes
                sMacroVersion = "2.2"
                sBuildSubVersion = "1"
                Call UpGradeSecurity2_1to2_2_1
            'Case vbNo
                'Call ExitMACRO
                'Call MACROEnd
        'End Select
    End If
   
    If sMacroVersion <> "2.2" Then
        MsgBox ("Your MACRO Security database is not valid. MACRO is being closed down.")
        ExitMACRO
        MACROEnd
    End If
    
    'ADD NEW 2.2 UPGRADES HERE
    If sBuildSubVersion = "1" Then
        sBuildSubVersion = "2"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If

    'REM 26/10/01 - 2 to 3
    If sBuildSubVersion = "2" Then
        sBuildSubVersion = "3"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If

    'REM 20/11/01 - 3 to 4
    If sBuildSubVersion = "3" Then
        sBuildSubVersion = "4"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'REM 10/12/01 - 4 to 5
    If sBuildSubVersion = "4" Then
        sBuildSubVersion = "5"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'MLM 20/12/01 - 5 to 6
    If sBuildSubVersion = "5" Then
        sBuildSubVersion = "6"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If

    'REM 11/01/02 - 6 to 7
    If sBuildSubVersion = "6" Then
        sBuildSubVersion = "7"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'REM 18/01/02 - 7 to 8
    If sBuildSubVersion = "7" Then
        sBuildSubVersion = "8"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
        Call UpgradeSecurity2_2From7to8
    End If

    'REM 31/01/02 - 8 to 9
    If sBuildSubVersion = "8" Then
        sBuildSubVersion = "9"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If

    'REM 15/04/02 - 9 to 10
    If sBuildSubVersion = "9" Then
        sBuildSubVersion = "10"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
        Call UpgradeSecurity2_2From9to10
    End If

    'MLM 28/04/02 - 10 to 11
    If sBuildSubVersion = "10" Then
        sBuildSubVersion = "11"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If

    'MLM 28/04/02 - 11 to 12
    If sBuildSubVersion = "11" Then
        sBuildSubVersion = "12"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'REM 10/05/02 - 12 to 13
    If sBuildSubVersion = "12" Then
        sBuildSubVersion = "13"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
        Call UpgradeSecurity2_2From12to13
    End If
    
    'ZA 14/05/02 - 13 to 14
    If sBuildSubVersion = "13" Then
        sBuildSubVersion = "14"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
        Call UpgradeSecurity2_2From13to14
    End If
    
    'REM 10/06/02 - 14 to 15
    If sBuildSubVersion = "14" Then
        sBuildSubVersion = "15"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If
    
    'Mo 11/06/02 - 15 to 16
    'MLM 25/06/02: Modified from 16 to 18 since 16 & 17 were used for patches
    If sBuildSubVersion = "15" Then
        sBuildSubVersion = "18"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
        Call UpgradeSecurity2_2From15to16
    End If
    
    'MLM 28/04/02 - 18 to 19
    If sBuildSubVersion = "18" Then
        sBuildSubVersion = "19"
        Call UpGradeSecurityToSubVersion(sBuildSubVersion)
    End If

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeSecurityDatabase2_2", "modUpgradeDatabases")
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
Private Sub UpGradeSecurity2_1to2_2_1()
'---------------------------------------------------------------------
' This upgrades a 2.1.x (latest) Security database to 2.2.1
'---------------------------------------------------------------------
Dim sSQL As String
Dim oCat As Catalog

    On Error GoTo ErrHandler
  
    ' *** Drop Columns ***
    sSQL = "ALTER Table Databases DROP Column DocumentLocation"
    moSecurityCon.Execute sSQL
    sSQL = "ALTER Table Databases DROP Column InFolderLocation"
    moSecurityCon.Execute sSQL
    sSQL = "ALTER Table Databases DROP Column OutFolderLocation"
    moSecurityCon.Execute sSQL
    sSQL = "ALTER Table Databases DROP Column TempLocation"
    moSecurityCon.Execute sSQL
    'drop index on field SubSystem prior to dropping the column
    sSQL = "ALTER Table Function DROP CONSTRAINT idx_SubSystem"
    moSecurityCon.Execute sSQL
    sSQL = "ALTER Table Function DROP Column SubSystem"
    moSecurityCon.Execute sSQL
    'drop index on field ClinicalTrialId prior to dropping the column
    sSQL = "ALTER Table UserRole DROP CONSTRAINT idx_ClinicalTrialId "
    moSecurityCon.Execute sSQL
    sSQL = "ALTER Table UserRole DROP Column ClinicalTrialId"
    moSecurityCon.Execute sSQL
    sSQL = "ALTER Table UserRole DROP Column TrialSite"
    moSecurityCon.Execute sSQL
    
    ' *** Foreign Keys ***
    'The following calls to FieldNameTypeUpgrade upgrade tables by first dropping them and then re-creating them.
    'Unfortunately the security database contains tables with foreign keys that prevent tables being dropped.
    'To get over this problem the foreign keys will be removed prior to the calls to FieldNameTypeUpgrade.
    'Note that FieldNameTypeUpgrade calls CreateDB to re-create the tables, with the foreign keys back in place.
    sSQL = "ALTER Table UserDatabase DROP CONSTRAINT FKUserdatabaseDD"
    moSecurityCon.Execute sSQL
    sSQL = "ALTER Table UserRole DROP CONSTRAINT FKUserRoleDD"
    moSecurityCon.Execute sSQL
    sSQL = "ALTER Table UserDatabase DROP CONSTRAINT FKUserdatabaseUC"
    moSecurityCon.Execute sSQL
    sSQL = "ALTER Table UserRole DROP CONSTRAINT FKUserRoleUC"
    moSecurityCon.Execute sSQL
    
    ' *** Columns with Name and/or Type Changes ***
    'Table Databases, field DatabaseDescription (Text(15)) changed to DatabaseCode (Text(15))
    Call FieldNameTypeUpgrade(eMACRODBFunction.Security, meDatabaseType, moSecurityCon, "Databases", "2.2.1")
    
    'Table MACROUser, field Password (Text(50)) changed to UserPassword (Text(50))
    'Table MACROUser, field UserCode (Text(50)) changed to UserName (Text(20))
    'Table MACROUser, field UserName (Text(50)) changed to UserNameFull (Text(100))
    Call FieldNameTypeUpgrade(eMACRODBFunction.Security, meDatabaseType, moSecurityCon, "MACROUser", "2.2.1")
    
    'Table UserDatabase, field DatabaseDescription (Text(15)) changed to DatabaseCode (Text(15))
    'Table UserDatabase, field UserCode(Text(50)) changed to UserName (Text(20))
    Call FieldNameTypeUpgrade(eMACRODBFunction.Security, meDatabaseType, moSecurityCon, "UserDatabase", "2.2.1")
    
    'Table UserRole, field DatabaseDescription (Text(15)) changed to DatabaseCode (Text(15))
    'Table UserRole, field UserCode(Text(50)) changed to UserName (Text(20))
    Call FieldNameTypeUpgrade(eMACRODBFunction.Security, meDatabaseType, moSecurityCon, "UserRole", "2.2.1")
    
    'Insert New Functions
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F5021','Remove own locks')"
    moSecurityCon.Execute sSQL

    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F5021')"
    moSecurityCon.Execute sSQL
    
    sSQL = "Insert INTO Function (FunctionCode,Function) "
    sSQL = sSQL & "VALUES ('F5022','Remove all locks')"
    moSecurityCon.Execute sSQL

    sSQL = "Insert INTO Rolefunction (RoleCode,FunctionCode) "
    sSQL = sSQL & "VALUES ('MacroUser','F5022')"
    moSecurityCon.Execute sSQL
    
    'Upgrade MACROVersion from [2.1] to [2.2]
    sSQL = "UPDATE SecurityControl Set MACROVersion = '2.2'"
    moSecurityCon.Execute sSQL

    'Upgrade BuildSubVersion to 1
    sSQL = "UPDATE SecurityControl Set BuildSubVersion = '1'"
    moSecurityCon.Execute sSQL

'    'change password column to UserPassword in MACROUser table
'    Set oCat = New Catalog
'    Set oCat.ActiveConnection = moSecurityCon
'    oCat.Tables("MACROUser").Columns("Password").Name = "UserPassword"
'    Set oCat = Nothing

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpGradeSecurity2_1to2_2_1", "modUpgradeDatabases")
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
Private Sub UpgradeSecurity2_2From7to8()
'---------------------------------------------------------------------
' REM 18/01/02
' Change the online support password to an encrypted one
'---------------------------------------------------------------------
Dim rsOnlineSupport As ADODB.Recordset

    On Error GoTo ErrHandler

    Set rsOnlineSupport = New ADODB.Recordset
    rsOnlineSupport.Open "Select SupportUserPassword from OnlineSupport where SupportUserName = 'INFERMED'" _
                , moSecurityCon, adOpenForwardOnly, adLockPessimistic, adCmdText
    rsOnlineSupport.Fields(0).Value = Crypt("guido")
    rsOnlineSupport.Update
    rsOnlineSupport.Close
    Set rsOnlineSupport = Nothing
    
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpgradeSecurity2_2From7to8", "modUpgradeDatabases")
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
Private Sub UpgradeSecurity2_2From9to10()
'---------------------------------------------------------------------
' REM 15/04/02
' Add new column to Security database Databases table called SecureHTMLLocation
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    'add column to Databases table
    sSQL = "ALTER Table Databases ADD COLUMN SecureHTMLLocation Text(255)"
    moSecurityCon.Execute sSQL
    
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpgradeSecurity2_2From9to10", "modUpgradeDatabases")
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
Private Sub UpgradeSecurity2_2From12to13()
'---------------------------------------------------------------------
'REM 10/05/02
'Change the permission called "Update Arezzo from GGB import" to
'"Update Arezzo from Clinical Gateway import"
'---------------------------------------------------------------------
Dim sSQL As String
    
    On Error GoTo ErrHandler
    
    sSQL = "UPDATE Function SET Function = 'Update Arezzo from Clinical Gateway import'" _
    & " WHERE FunctionCode = 'F3023'"
    moSecurityCon.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpgradeSecurity2_2From12to13", "modUpgradeDatabases")
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
Private Sub UpgradeSecurity2_2From13to14()
'---------------------------------------------------------------------
'ZA 14/05/02
'delete F3015 (create report) and F5006 (view reports) function
'from Function and RoleFunction table
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler
    
    sSQL = "Delete from RoleFunction where FunctionCode ='F3015'"
    moSecurityCon.Execute sSQL
    
    sSQL = "Delete from RoleFunction where FunctionCode ='F5006'"
    moSecurityCon.Execute sSQL
    
    sSQL = "Delete from Function where FunctionCode ='F3015'"
    moSecurityCon.Execute sSQL
    
    sSQL = "Delete from Function where FunctionCode ='F5006'"
    moSecurityCon.Execute sSQL
    
Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, _
                                "UpgradeSecurity2_2From13to14", "modUpgradeDatabases")
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
Private Sub UpgradeSecurity2_2From15to16()
'---------------------------------------------------------------------
'Mo Morris 11/6/2002, Add new Function code 'F1008' to table Function
'   and table Rolefunction for RoleCode 'MacroUser'
'---------------------------------------------------------------------
Dim sSQL As String

    On Error GoTo ErrHandler

    sSQL = "INSERT INTO Function (FunctionCode,Function) VALUES ('F1008','Access Query Module')"
    moSecurityCon.Execute sSQL
    sSQL = "INSERT INTO Rolefunction (RoleCode,FunctionCode) VALUES ('MacroUser','F1008')"
    moSecurityCon.Execute sSQL
    
'TA 13/06/02 CBB 2.2.13.43 F5006 (view reports) function reinserted into Function and RoleFunction tables
    sSQL = "Insert into Function values ('F5006','View reports')"
    moSecurityCon.Execute sSQL
    
    sSQL = "insert into RoleFunction values ('MacroUser','F5006')"
    moSecurityCon.Execute sSQL

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "UpgradeSecurity2_2From15to16", "modUpgradeDatabases.bas")
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
Private Sub FieldNameTypeUpgrade(ByVal nDBFunction As eMACRODBFunction, _
                                ByVal nDBType As MACRODatabaseType, _
                                ByVal oConnection As ADODB.Connection, _
                                ByVal sTable As String, _
                                ByVal sVersion As String)
'---------------------------------------------------------------------
'Utility used to upgrade a table with one or more fields that have had
'a Name and/or Type change.
'
'sTable contains the table to be upgraded.
'sVersion contains the version that the table is being updated to.
'
'The upgrade happens in 4 stages:-
'   Read the contents of the specified table out into a recordset.
'   Drop the old version of the table.
'   Create a new version of the table using CreateDB.
'   Read the contents of the recordset back into the newly created version of the table.
'
'Note that if a table has had new fields added, or old fields removed, at the same time
'as other fields have had name/type changes then the adding and/or dropping should take place
'prior to this sub being called. (i.e. the old and new versions of a table need to have the
'same number of fields for this sub to work)
'
'THIS UTILITY WILL HANDLE TYPE CHANGES BETWEEN INTEGER AND LONG
'THIS UTILITY WILL HANDLE TYPE CHANGES BETWEEN AND STRINGS OF DIFFERENT LENGTHS
'THIS UTILITY WILL NOT HANDLE TYPE CHANGES THAT REQUIRE A FUNDAMENTAL TYPE CHANGE (I.E. INTEGER TO STRING)
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsFromTable As ADODB.Recordset
Dim rsToTable As ADODB.Recordset
Dim j As Long
Dim i As Integer
Dim lNumberOfRecords As Long
Dim nNumberOfFields As Integer

    On Error GoTo ErrHandler

    'Place the contents of the table into recordset rsFromTable
    sSQL = "SELECT * FROM " & sTable
    Set rsFromTable = New ADODB.Recordset
    rsFromTable.CursorLocation = adUseClient
    rsFromTable.Open sSQL, oConnection, adOpenKeyset, adLockReadOnly, adCmdText
    rsFromTable.ActiveConnection = Nothing
    
    'Drop the original version of the table
    sSQL = "DROP Table " & sTable
    oConnection.Execute sSQL
    
    'Create a new version of the table
    'Mo Morris 1/10/01, optional field to prevent CreateDB displaying messages added
    Call CreateDB(nDBFunction, nDBType, oConnection, False, False, sTable, sVersion)
    
    'Prepare recordset rsToTable to receive the contents of rsFromTable
    sSQL = "SELECT * FROM " & sTable
    Set rsToTable = New ADODB.Recordset
    rsToTable.Open sSQL, oConnection, adOpenKeyset, adLockPessimistic, adCmdText
    
    'copy the contents of rsTemp into the newly created table
    lNumberOfRecords = rsFromTable.RecordCount
    nNumberOfFields = rsFromTable.Fields.Count
    For j = 1 To lNumberOfRecords
        rsToTable.AddNew
        For i = 0 To nNumberOfFields - 1
            rsToTable.Fields(i).Value = rsFromTable.Fields(i).Value
        Next i
        rsToTable.Update
        rsFromTable.MoveNext
    Next j
    
    rsToTable.Close
    Set rsToTable = Nothing
    rsFromTable.Close
    Set rsFromTable = Nothing

Exit Sub
ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "FieldNameTypeUpgrade", "modUpgradeDatabases.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select

End Sub

'--------------------------------------------------------------------------------
Private Function CallVersionLessThanTagVersion(ByVal sCallingVersion As String, _
                                                ByVal sTagVersion As String) As Boolean
'--------------------------------------------------------------------------------
Dim asCallingVersion() As String
Dim asTagVersion() As String
Dim lCallingVersion As Long
Dim lTagVersion As Long

    On Error GoTo ErrHandler

    asCallingVersion = Split(sCallingVersion, ".")
    asTagVersion = Split(sTagVersion, ".")
    lCallingVersion = Val(Format(asCallingVersion(0), "00") & Format(asCallingVersion(1), "00") & Format(asCallingVersion(2), "0000"))
    lTagVersion = Val(Format(asTagVersion(0), "00") & Format(asTagVersion(1), "00") & Format(asTagVersion(2), "0000"))
    
    If lCallingVersion < lTagVersion Then
        CallVersionLessThanTagVersion = True
    Else
        CallVersionLessThanTagVersion = False
    End If
    
Exit Function

ErrHandler:
    Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "CallVersionLessThanTagVersion", "modNewDatabasesAndTables")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Function

