VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------
' File: UserDetail.cls
' Copyright: InferMed Ltd 2004 All Rights Reserved
' Author: Toby Aldridge, September 2004
' Purpose: Store the properties of a suer to be changed by web client
'------------------------------------------------------------------
' REVISIONS

'------------------------------------------------------------------

Option Explicit

Private msUserName As String
Private msUserNameFull As String
Public Enabled As Boolean
Public FailedAttempts As Integer
Public UnEncryptedPassword As String
Private mbSysAdmin As Boolean

'------------------------------------------------------------------
Public Property Get UserName() As String
'------------------------------------------------------------------

    UserName = msUserName
End Property

'------------------------------------------------------------------
Public Property Let UserName(sUserName As String)
'------------------------------------------------------------------
    msUserName = sUserName
End Property

'------------------------------------------------------------------
Public Property Get UserNameFull() As String
'------------------------------------------------------------------
    UserNameFull = msUserNameFull
End Property

'------------------------------------------------------------------
Friend Property Let UserNameFull(sUserNameFull As String)
'------------------------------------------------------------------
    msUserNameFull = sUserNameFull
End Property

'------------------------------------------------------------------
Public Property Get SysAdmin() As Boolean
'------------------------------------------------------------------
    SysAdmin = mbSysAdmin
End Property

'------------------------------------------------------------------
Friend Property Let SysAdmin(bSysAdmin As Boolean)
'------------------------------------------------------------------
    mbSysAdmin = bSysAdmin
End Property

Public Function ToString() As String
'------------------------------------------------------------------
' give string representation of object
'------------------------------------------------------------------
    ToString = "Username:" & msUserName & ", " _
        & "UsernameFull:" & msUserNameFull & ", " _
        & "Enabled:" & Enabled & ", " _
        & "FailedAttempts:" & FailedAttempts & ", " _
        & "SysAdmin:" & SysAdmin & ", " _
        & "Password: not telling"
        
End Function

