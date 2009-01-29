VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   File:       UserSettings.cls
'   Author:     Toby Aldridge, November 2002
'   Purpose:    Set, retrieve usersettings
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   NCJ 27 Nov 03 - Tidy up on termination
'   DPH 24/03/2004 - Clean up collections in class_terminate
'
'----------------------------------------------------------------------------------------'

Option Explicit

'Public SplitScreen As Boolean
'Public ViewSymbols As Boolean
'Public ViewFunctions As Boolean
'Public LocalFormat As Boolean
'Public ServerTime As Boolean
'Public RecordsPerPage As Long
'Public SameForm As Boolean

Private msUserName As String
Private msCon As String

Private mcolSettings As Dictionary

'----------------------------------------------------------------------------------------'
Public Function GetSetting(sSettingName As String, vDefault As Variant) As Variant
'----------------------------------------------------------------------------------------'
    
    If mcolSettings.Exists(sSettingName) Then
        'TA 07/04/2003: convert is needed
        GetSetting = ConvertSettingFromString(mcolSettings.Item(sSettingName))
    Else
        GetSetting = vDefault
    End If
    
Exit Function

ErrLabel:
    
    
    
End Function

'----------------------------------------------------------------------------------------'
Public Sub SetSetting(sSetting As String, vValue As Variant)
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'
Dim oCon As Connection
Dim lRecords As Long
Dim sUpdateSQL As String
Dim sInsertSQL As String

    On Error Resume Next
    mcolSettings.Remove sSetting
    Err.Clear
    On Error GoTo ErrLabel
    
    mcolSettings.Add sSetting, vValue
    
    sUpdateSQL = "UPDATE MACROUSERSETTINGS SET SETTINGVALUE = '" & ConvertSettingToString(vValue) & "'"
    sUpdateSQL = sUpdateSQL & " WHERE USERNAME = '" & msUserName & "'"
    sUpdateSQL = sUpdateSQL & " AND USERSETTING = '" & sSetting & "'"
    
    sInsertSQL = "INSERT INTO MACROUSERSETTINGS VALUES ("
    sInsertSQL = sInsertSQL & "'" & msUserName & "','" & sSetting & "','" & ConvertSettingToString(vValue) & "')"
    
    Set oCon = New Connection
    oCon.Open msCon
    oCon.CursorLocation = adUseClient
    oCon.Execute sUpdateSQL, lRecords
    
    If lRecords = 0 Then
        'no rows updated, let's insert
        oCon.Execute sInsertSQL
    End If
    
    oCon.Close
    Set oCon = Nothing
    
    Exit Sub
ErrLabel:
   Err.Raise Err.Description, , Err.Description & "|UserSettings.SetSetting"
   
End Sub

'----------------------------------------------------------------------------------------'
Public Sub Load(sCon As String, sUserName As String)
'----------------------------------------------------------------------------------------'
'load in the settings
'----------------------------------------------------------------------------------------'
Dim oCon As Connection
Dim rs As Recordset


On Error GoTo ErrLabel
'skip out on errors (table does nto exist yet)
    msCon = sCon
    msUserName = sUserName
'   set defaults
    Set mcolSettings = New Dictionary
    Set oCon = New Connection
    Set rs = New Recordset
    oCon.Open msCon
    oCon.CursorLocation = adUseClient
    rs.Open "SELECT USERSETTING, SETTINGVALUE FROM MACROUSERSETTINGS WHERE USERNAME = '" & msUserName & "'", oCon
    
    If rs.RecordCount <> 0 Then
        Do While Not rs.EOF
            mcolSettings.Add rs.Fields(0).Value, rs.Fields(1).Value
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    Set rs = Nothing
    oCon.Close
    Set oCon = Nothing
    Exit Sub
ErrLabel:
    

End Sub

'----------------------------------------------------------------------------------------'
Private Function ConvertSettingToString(vValue As Variant) As String
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    Select Case VarType(vValue)
    Case vbBoolean
        If vValue Then
            ConvertSettingToString = "true"
        Else
            ConvertSettingToString = "false"
        End If
    Case vbString
        ConvertSettingToString = vValue
    Case Else
        'numeric
        ConvertSettingToString = vValue
    End Select
    
End Function



'----------------------------------------------------------------------------------------'
Private Function ConvertSettingFromString(sDBValue As String) As Variant
'----------------------------------------------------------------------------------------'

'----------------------------------------------------------------------------------------'

    Select Case LCase(sDBValue)
    Case "true"
        ConvertSettingFromString = True
    Case "false"
        ConvertSettingFromString = False
    Case Else
        ConvertSettingFromString = sDBValue
    End Select

    
End Function


'----------------------------------------------------------------------------------------'
Private Sub Class_ReadProperties(PropBag As PropertyBag)
'----------------------------------------------------------------------------------------'
'support for serialisation
'----------------------------------------------------------------------------------------'
Dim i As Long

    msUserName = PropBag.ReadProperty("u")
    msCon = PropBag.ReadProperty("c")
    Set mcolSettings = New Dictionary
    For i = 0 To PropBag.ReadProperty("co") - 1
        mcolSettings.Add Split(PropBag.ReadProperty("s" & i), ",")(0), Split(PropBag.ReadProperty("s" & i), ",")(1)
    Next
    
End Sub

'----------------------------------------------------------------------
Private Sub Class_Terminate()
'----------------------------------------------------------------------
' NCJ 27 Nov 03 - Tidy up on termination
' DPH 24/03/2004 - remove all from dictionary
'----------------------------------------------------------------------

    If (Not mcolSettings Is Nothing) Then
        mcolSettings.RemoveAll
    End If

    Set mcolSettings = Nothing

End Sub

'----------------------------------------------------------------------------------------'
Private Sub Class_WriteProperties(PropBag As PropertyBag)
'----------------------------------------------------------------------------------------'
'support for serialisation
'----------------------------------------------------------------------------------------'
Dim i As Long

     PropBag.WriteProperty "u", msUserName
    PropBag.WriteProperty "c", msCon
    PropBag.WriteProperty "co", mcolSettings.Count
    For i = 0 To mcolSettings.Count - 1
        PropBag.WriteProperty "s" & i, mcolSettings.Keys(i) & "," & mcolSettings.Items(i)
    Next


End Sub
