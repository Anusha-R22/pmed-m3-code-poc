VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SysMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999-2002. All Rights Reserved
'   File:       SysMessages.cls
'   Author      Richard Meinesz, 2002
'   Purpose:    Writes system messages to the message table
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'
'----------------------------------------------------------------------------------------'

Option Explicit


'----------------------------------------------------------------------------------------'
Public Sub AddNewSystemMessage(conMACRO As ADODB.Connection, nMessageType As Integer, sSystemUserName As String, sUserName As String, _
                            sMessageBody As String, sMessageParameters As String, Optional sSite As String = "", Optional sRoleCode As String)
'----------------------------------------------------------------------------------------'
'REM 18/11/02
'Adds a system message to the MACRO Message table
'----------------------------------------------------------------------------------------'
    
    On Error GoTo ErrLabel
    
    Call AddSystemMessage(conMACRO, nMessageType, sSystemUserName, sUserName, sMessageBody, sMessageParameters, sSite, sRoleCode)
    
Exit Sub
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "clsSystemMessages.AddNewSystemMessage"
End Sub

'----------------------------------------------------------------------------------------'
Public Function GetSiteCode(conMACRO As ADODB.Connection, sSettingSection As String, sSettingKey As String, _
                            sDefault As String) As String
'----------------------------------------------------------------------------------------'
'REM 18/11/02
'Returns the site code of a database if it is a site database, otherwise will return the string values et as sDefault
'----------------------------------------------------------------------------------------'

    On Error GoTo ErrLabel
    
    GetSiteCode = GetDBSettings(conMACRO, sSettingSection, sSettingKey, sDefault)

Exit Function
ErrLabel:
Err.Raise Err.Number, , Err.Description & "|" & "clsSystemMessages.GetSiteCode"
End Function

