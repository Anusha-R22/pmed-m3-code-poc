VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChangeUserDetailsTest"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------
' File: ChangeUserDetailsTest.cls
' Copyright: InferMed Ltd 2004 All Rights Reserved
' Author: Toby Aldridge, September 2004
' Purpose: TESITNG PURPOSES ONLY
'------------------------------------------------------------------
' REVISIONS

'------------------------------------------------------------------
Option Explicit

Public Function TestGetUserDetail()

Dim oAPI As MACROAPI
Dim sSerialisedUser As String
Dim sMessage As String
Dim sUserNameFull As String

    Set oAPI = New MACROAPI
    
    Debug.Print "logn result - " & oAPI.Login("t", "a", "ora30ta2", "MACROUser", sMessage, sUserNameFull, sSerialisedUser)
    Debug.Print "login smessage = " & sMessage
    
    Dim colUsers As New Collection
    Dim oDetail As UserDetail
    
    Set colUsers = oAPI.GetUsersDetails(sSerialisedUser, "", sMessage)
    Debug.Print "getusersdetails smessage " & sMessage
    
    For Each oDetail In colUsers
        Debug.Print oDetail.ToString
    Next
    Debug.Print "_______________________________________________________________________"
End Function


Public Function TestGetSIngleUserDetail()

Dim oAPI As MACROAPI
Dim sSerialisedUser As String
Dim sMessage As String
Dim sUserNameFull As String

    Set oAPI = New MACROAPI
    
    Debug.Print "logn result - " & oAPI.Login("t", "a", "ora30ta2", "MACROUser", sMessage, sUserNameFull, sSerialisedUser)
    Debug.Print "login smessage = " & sMessage
    
    Dim colUsers As New Collection
    Dim oDetail As UserDetail
    
    Set colUsers = oAPI.GetUsersDetails(sSerialisedUser, "rick", sMessage)
    Debug.Print "getusersdetails smessage " & sMessage
    
    For Each oDetail In colUsers
        Debug.Print oDetail.ToString
    Next
    Debug.Print "_______________________________________________________________________"
End Function

Public Function TestUpdateUserDetail()

Dim oAPI As MACROAPI
Dim sSerialisedUser As String
Dim sMessage As String
Dim sUserNameFull As String
Dim oDetail As New UserDetail

    With oDetail
        .UserName = "ta2"
        .Enabled = True
        .FailedAttempts = 4
    End With
    Set oAPI = New MACROAPI
    
    Debug.Print "logn result - " & oAPI.Login("t", "a", "ora30ta2", "MACROUser", sMessage, sUserNameFull, sSerialisedUser)
    Debug.Print "login smessage = " & sMessage

    Debug.Print oAPI.ChangeUserDetails(sSerialisedUser, oDetail, sMessage)
    
    Debug.Print "change user details smessage: " & sMessage
Debug.Print "_______________________________________________________________________"
End Function


Public Function TestUpdateUserPassword()

Dim oAPI As MACROAPI
Dim sSerialisedUser As String
Dim sMessage As String
Dim sUserNameFull As String
Dim oDetail As New UserDetail

    With oDetail
        .UserName = "ta2"
        .UnEncryptedPassword = "toby"
        '.FailedAttempts = 1
    End With
    Set oAPI = New MACROAPI
    
    Debug.Print "logn result - " & oAPI.Login("t", "a", "ora30ta2", "MACROUser", sMessage, sUserNameFull, sSerialisedUser)
    Debug.Print "login smessage = " & sMessage

    Debug.Print oAPI.ChangeUserDetails(sSerialisedUser, oDetail, sMessage)
    
    Debug.Print "change user details smessage" & sMessage
Debug.Print "_______________________________________________________________________"
End Function
