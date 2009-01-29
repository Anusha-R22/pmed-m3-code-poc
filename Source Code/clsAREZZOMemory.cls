VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAREZZOMemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------
'   File:       clsAREZZOMemory.cls
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   Author:     Ashitei Trebi-Ollennu, November 2002
'   Purpose:    AREZZO Memory Settings
'----------------------------------------------------------------------------------------
'Revisions:
'   NCJ 17 Jan 03 - Calling Load with TrialId = -1 returns default settings
'   NCJ 28 Jan 03 - Pass in DB connection string for Load and Save
'                   (so we can use class in Subject Cache)
'           Also allow switch values in the Settings file to override the DB values
'----------------------------------------------------------------------------------------

Option Explicit

Private mlProgramSpace As Long
Private mlTextSpace As Long
Private mlLocalSpace As Long
Private mlBacktrackSpace As Long
Private mlHeapSpace As Long
Private mlInputSpace As Long
Private mlOutputSpace As Long
Private mlTrialID As Long

'--------------------------------------------------------------------------
Public Property Get ProgramSpace() As Long
'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------

    ProgramSpace = mlProgramSpace

End Property

'--------------------------------------------------------------------------
Public Property Let ProgramSpace(lValue As Long)
'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
    
    mlProgramSpace = lValue

End Property

'--------------------------------------------------------------------------
Public Property Get TextSpace() As Long
'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------

    TextSpace = mlTextSpace

End Property

'---------------------------------------------------------------------------
Public Property Let TextSpace(lValue As Long)
'---------------------------------------------------------------------------
'
'---------------------------------------------------------------------------
    
    mlTextSpace = lValue

End Property

'---------------------------------------------------------------------------
Public Property Get LocalSpace() As Long
'---------------------------------------------------------------------------
'
'---------------------------------------------------------------------------

    LocalSpace = mlLocalSpace

End Property

'----------------------------------------------------------------------------
Public Property Let LocalSpace(lValue As Long)
'----------------------------------------------------------------------------
'
'----------------------------------------------------------------------------

     mlLocalSpace = lValue

End Property

'----------------------------------------------------------------------------
Public Property Get BacktrackSpace() As Long
'----------------------------------------------------------------------------
'
'----------------------------------------------------------------------------

    BacktrackSpace = mlBacktrackSpace

End Property

'----------------------------------------------------------------------------
Public Property Let BacktrackSpace(lValue As Long)
'----------------------------------------------------------------------------
'
'----------------------------------------------------------------------------

    mlBacktrackSpace = lValue
     
End Property

'---------------------------------------------------------------------------
Public Property Get HeapSpace() As Long
'---------------------------------------------------------------------------
'
'---------------------------------------------------------------------------

    HeapSpace = mlHeapSpace

End Property

'---------------------------------------------------------------------------
Public Property Let HeapSpace(lValue As Long)
'---------------------------------------------------------------------------
'
'---------------------------------------------------------------------------

   mlHeapSpace = lValue

End Property

'---------------------------------------------------------------------------
Public Property Get InputSpace() As Long
'---------------------------------------------------------------------------
'
'---------------------------------------------------------------------------

    InputSpace = mlInputSpace

End Property

'---------------------------------------------------------------------------
Public Property Let InputSpace(lValue As Long)
'---------------------------------------------------------------------------
'
'---------------------------------------------------------------------------

    mlInputSpace = lValue

End Property

'---------------------------------------------------------------------------
Public Property Get OutputSpace() As Long
'---------------------------------------------------------------------------
'
'---------------------------------------------------------------------------

    OutputSpace = mlOutputSpace

End Property

'----------------------------------------------------------------------------
Public Property Let OutputSpace(lValue As Long)
'----------------------------------------------------------------------------
'
'----------------------------------------------------------------------------

     mlOutputSpace = lValue

End Property

'-----------------------------------------------------------------------------
Public Property Get AREZZOSwitches(Optional bUseSettingsFileValues As Boolean = True) As String
'-----------------------------------------------------------------------------
' Creates a string of AREZZO Memory settings based on the current values
' NCJ 28 Jan 03 - If bUseSettingsFileValues = TRUE, values in the Settings file will override the current values
'-----------------------------------------------------------------------------
Dim sAREZZOSettings As String
    
    On Error GoTo ErrHandler
    
    sAREZZOSettings = ""
    sAREZZOSettings = sAREZZOSettings & "/P" & GetSwitchValue(gsPROGRAM_SPACE, mlProgramSpace, bUseSettingsFileValues)
    sAREZZOSettings = sAREZZOSettings & " /T" & GetSwitchValue(gsTEXT_SPACE, mlTextSpace, bUseSettingsFileValues)
    sAREZZOSettings = sAREZZOSettings & " /L" & GetSwitchValue(gsLOCAL_SPACE, mlLocalSpace, bUseSettingsFileValues)
    sAREZZOSettings = sAREZZOSettings & " /B" & GetSwitchValue(gsBACKTRACK_SPACE, mlBacktrackSpace, bUseSettingsFileValues)
    sAREZZOSettings = sAREZZOSettings & " /H" & GetSwitchValue(gsHEAP_SPACE, mlHeapSpace, bUseSettingsFileValues)
    sAREZZOSettings = sAREZZOSettings & " /I" & GetSwitchValue(gsINPUT_SPACE, mlInputSpace, bUseSettingsFileValues)
    sAREZZOSettings = sAREZZOSettings & " /O" & GetSwitchValue(gsOUTPUT_SPACE, mlOutputSpace, bUseSettingsFileValues)
    AREZZOSwitches = sAREZZOSettings
    
Exit Property
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|clsAREZZOMemory.AREZZOSwitches"
End Property

'------------------------------------------------------------------------------
Private Function GetSwitchValue(sKey As String, lValue As Long, bUseSettingsFileValues As Boolean) As String
'------------------------------------------------------------------------------
' Get a switch value for adding to the Switches string
' Proposed value is lValue,
' but if bUseSettingsFileValues = TRUE then we override with a value from the IMedSettings file if it exists
'------------------------------------------------------------------------------
    
    If bUseSettingsFileValues Then
        GetSwitchValue = GetMACROSetting(sKey, lValue)
    Else
        GetSwitchValue = lValue
    End If

End Function

'------------------------------------------------------------------------------
Public Sub SaveValues(sDBCon As String)
'------------------------------------------------------------------------------
'saves the setting values to the database
' NCJ 17 Jan 03 - In "default" mode we may not have a TrialId
' NCJ 28 Jan 03 - Added sDBCon as DB connection string
'------------------------------------------------------------------------------
Dim sSQL As String
Dim oDBCon As ADODB.Connection
    
    On Error GoTo ErrHandler
    
    If mlTrialID > 0 Then
        Set oDBCon = New ADODB.Connection
        oDBCon.Open sDBCon
    
        ' NCJ 28 Jan 03 - Get switches but DO NOT override with Settings file values
        sSQL = "UPDATE StudyDefinition SET ArezzoMemory ='" & AREZZOSwitches(False) & "'"
        sSQL = sSQL & " WHERE ClinicalTrialID = " & mlTrialID
        oDBCon.Execute sSQL
        
        oDBCon.Close
        Set oDBCon = Nothing
    End If
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|clsAREZZOMemory.SaveValues"
End Sub

'-------------------------------------------------------------------------------------
Public Sub Load(ByVal lTrialId As Long, sDBCon As String)
'-------------------------------------------------------------------------------------
'gets the value for each individual AREZZO memory settings in the database
'if trialid is = 0 then it loops through all studies in the database and finds the
'maximum for each
'if trialid is > 0 then it load for just that study
' NCJ 17 Jan 03 - If lTrialId = -1, load with default settings
' NCJ 28 Jan 03 - Takes sDBCon connection string (so we can call it from the Subject Cache)
'-------------------------------------------------------------------------------------
Dim sSQL As String
Dim rsArezzoMemory As ADODB.Recordset
Dim sAREZZOSettings As String
Dim sArezzoString As String
Dim var As Variant
Dim n As Integer
Dim nCounter As Integer
Dim lVal As Long
Dim oDBCon As ADODB.Connection

    On Error GoTo ErrHandler
    
    'set setting value to default
    InitialiseDefaultValues
    
    mlTrialID = lTrialId
    
    ' NCJ 17 Jan 03 - If we just want default settings, nothing more to do
    If lTrialId = -1 Then Exit Sub
    
    Set oDBCon = New ADODB.Connection
    oDBCon.Open sDBCon
    oDBCon.CursorLocation = adUseClient     ' NCJ 11 Feb 03
    
    If mlTrialID = 0 Then
         sSQL = "SELECT AREZZOMemory from StudyDefinition WHERE ClinicaltrialID > 0"
    Else
        sSQL = "SELECT AREZZOMemory from StudyDefinition WHERE ClinicaltrialID = " & lTrialId
    End If
    
    Set rsArezzoMemory = New ADODB.Recordset
    rsArezzoMemory.Open sSQL, oDBCon, adOpenKeyset, adLockOptimistic, adCmdText
    
    If rsArezzoMemory.RecordCount > 0 Then
        rsArezzoMemory.MoveFirst
        Do Until rsArezzoMemory.EOF
            sAREZZOSettings = ConvertFromNull(rsArezzoMemory.Fields(0).Value, vbString)
            If sAREZZOSettings <> "" Then
                'if there are some settings stored then unwrap the string
                var = Split(sAREZZOSettings, "/")
                'compare intial value to new value
                'and replace with the higher value each time
                For n = 1 To UBound(var)
                    lVal = Mid(var(n), 2)
                    Select Case Left(var(n), 1)
                    Case "P"
                        If lVal > mlProgramSpace Then
                            mlProgramSpace = lVal
                        End If
                    Case "T"
                        If lVal > mlTextSpace Then
                            mlTextSpace = lVal
                        End If
                    Case "L"
                        If lVal > mlLocalSpace Then
                            mlLocalSpace = lVal
                        End If
                    Case "B"
                        If lVal > mlBacktrackSpace Then
                            mlBacktrackSpace = lVal
                        End If
                    Case "H"
                        If lVal > mlHeapSpace Then
                            mlHeapSpace = lVal
                        End If
                    Case "I"
                        If lVal > mlInputSpace Then
                            mlInputSpace = lVal
                        End If
                    Case "O"
                        If lVal > mlOutputSpace Then
                            mlOutputSpace = lVal
                        End If
                    End Select
                Next
            End If
            rsArezzoMemory.MoveNext
        Loop
    End If

    rsArezzoMemory.Close
    Set rsArezzoMemory = Nothing
    oDBCon.Close
    Set oDBCon = Nothing
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|clsAREZZOMemory.Load"
End Sub

'--------------------------------------------------------------------------------
Private Sub InitialiseDefaultValues()
'--------------------------------------------------------------------------------
'loads default AREZZO memory settings values
'--------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    mlProgramSpace = glPROGRAM_SPACE
    mlTextSpace = glTEXT_SPACE
    mlLocalSpace = glLOCAL_SPACE
    mlBacktrackSpace = glBACKTRACK_SPACE
    mlHeapSpace = glHEAP_SPACE
    mlInputSpace = glINPUT_SPACE
    mlOutputSpace = glOUTPUT_SPACE

Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|clsAREZZOMemory.InitialiseDefaultValues"
End Sub

