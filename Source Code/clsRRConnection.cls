VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRRConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       clsRRConnection.cls
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Toby Aldridge, November 2000
'   Purpose:    Randomisation / Registration Connection object for MACRO
'----------------------------------------------------------------------------------------'
' REVISIONS:
' TA 06/12/2000: Minor code changes from white box test

Option Explicit

Public Event ValidHTTPAddress(bIsValid As Boolean)
Public Event ValidUserName(bIsValid As Boolean)
Public Event ValidPassword(bIsValid As Boolean)
Public Event ValidProxyServer(bIsValid As Boolean)

Public Event ConnectionDetailsChange(bAllow As Boolean)

Public Event IsValid(bValid As Boolean)
Public Event HasChanges(bHasChanges As Boolean)

Private mnServerType As Integer
Private msHTTPAddress As String
Private msUserName As String
Private msPassword As String
Private msProxyServer As String

Private mlClinicalTrialId As Long
Private mnVersionId As Integer

Private WithEvents moBrokenRules As clsBrokenRules
Attribute moBrokenRules.VB_VarHelpID = -1
Private mbHasChanges As Boolean
Private mbIsValid As Boolean

'----------------------------------------------------------------------------------------'
Public Sub Load(lClinicalTrialId As Long, nVersionId As Integer)
'----------------------------------------------------------------------------------------'
'load connection details
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim rsRR As ADODB.Recordset


    On Error GoTo ErrHandler
    
    mlClinicalTrialId = lClinicalTrialId
    mnVersionId = nVersionId
    sSQL = "SELECT RRServerType, RRHTTPAddress, RRUserName, RRPassword, RRProxyServer FROM StudyDefinition" _
                & " WHERE ClinicalTrialId = " & mlClinicalTrialId & " AND VersionId = " & mnVersionId
    
    Set rsRR = New ADODB.Recordset
    
    With rsRR
        .Open sSQL, MacroADODBConnection
        'one and only one row will always be returned
        mnServerType = Val(RemoveNull(.Fields("RRServerType")))
        msHTTPAddress = RemoveNull(.Fields("RRHTTPAddress"))
        msUserName = RemoveNull(.Fields("RRUsername"))
        msPassword = RemoveNull(.Fields("RRPassword"))
        msProxyServer = RemoveNull(.Fields("RRProxyServer"))
        .Close
    End With
    Set rsRR = Nothing
    
    Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "Load", "clsRRConnection")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
        
End Sub

'----------------------------------------------------------------------------------------'
Public Sub Save()
'----------------------------------------------------------------------------------------'
'save connection details
'----------------------------------------------------------------------------------------'
Dim sSQL As String

    On Error GoTo ErrHandler
    
    sSQL = "UPDATE StudyDefinition SET" _
            & " RRServerType = " & mnServerType _
            & ", RRHTTPAddress = " & StringToSQLString(msHTTPAddress) _
            & ", RRUserName = " & StringToSQLString(msUserName) _
            & ", RRPassword = " & StringToSQLString(msPassword) _
            & ", RRProxyServer = " & StringToSQLString(msProxyServer) _
            & " WHERE ClinicalTrialId = " & mlClinicalTrialId & " AND VersionId = " & mnVersionId

    MacroADODBConnection.Execute sSQL
    
Exit Sub

ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "Save", "clsRRConnection")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
   
End Sub

'----------------------------------------------------------------------------------------'
Public Property Let ServerType(nType As Integer)
'----------------------------------------------------------------------------------------'

    If mnServerType <> nType Then
        'has changed
        mnServerType = nType
        
        'tell form whether connection details can be edited
        RaiseEvent ConnectionDetailsChange(ConnectionDetailsChange)
        
        Call RRServerHasAddress
        
        HasChanges = True
    End If
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get ServerType() As Integer
'----------------------------------------------------------------------------------------'

    ServerType = mnServerType
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Let HTTPAddress(sAddress As String)
'----------------------------------------------------------------------------------------'

    If msHTTPAddress <> sAddress Then
        'has changed
        msHTTPAddress = sAddress
        
        Call RRServerHasAddress
    
        HasChanges = True
    End If
End Property

'----------------------------------------------------------------------------------------'
Public Property Get HTTPAddress() As String
'----------------------------------------------------------------------------------------'

    HTTPAddress = msHTTPAddress

End Property

'----------------------------------------------------------------------------------------'
Public Property Let UserName(sUsername As String)
'----------------------------------------------------------------------------------------'

    If msUserName <> sUsername Then
        'has changed
        msUserName = sUsername
         
        If Len(msUserName) <= 50 Then
            moBrokenRules.RuleBroken "Username", False
            'raise this event to inform the display that it is valid
            RaiseEvent ValidUserName(True)
            
        Else
            'raise this event to inform the display that it is invalid
            moBrokenRules.RuleBroken "Username", True
            RaiseEvent ValidUserName(False)
        End If
       
        HasChanges = True
    End If
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get UserName() As String
'----------------------------------------------------------------------------------------'

    UserName = msUserName

End Property

'----------------------------------------------------------------------------------------'
Public Property Let Password(sPassword As String)
'----------------------------------------------------------------------------------------'
' validation rules for settting password
'----------------------------------------------------------------------------------------'
    If msPassword <> sPassword Then
        'has changed
        msPassword = sPassword
        
        If Len(msPassword) <= 50 Then
            moBrokenRules.RuleBroken "Password", False
            'raise this event to inform the display that it is valid
            RaiseEvent ValidPassword(True)
            
        Else
            'raise this event to inform the display that it is invalid
            moBrokenRules.RuleBroken "Password", True
            RaiseEvent ValidPassword(False)
        End If
    
        HasChanges = True
    End If
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get Password() As String
'----------------------------------------------------------------------------------------'

    Password = msPassword

End Property

'----------------------------------------------------------------------------------------'
Public Property Let ProxyServer(sProxyServer As String)
'----------------------------------------------------------------------------------------'

    If msProxyServer <> sProxyServer Then
        'has changed
        msProxyServer = sProxyServer
        
        'check remote server has address rulke
        Call RRServerHasAddress
        
        HasChanges = True
    End If
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get ProxyServer() As String
'----------------------------------------------------------------------------------------'

    ProxyServer = msProxyServer

End Property

'----------------------------------------------------------------------------------------'
Public Property Let HasChanges(bHasChanges As Boolean)
'----------------------------------------------------------------------------------------'

    mbHasChanges = bHasChanges
    RaiseEvent HasChanges(mbHasChanges)

End Property

'----------------------------------------------------------------------------------------'
Public Property Get HasChanges() As Boolean
'----------------------------------------------------------------------------------------'

    HasChanges = mbHasChanges
    
End Property

'----------------------------------------------------------------------------------------'
Private Property Let IsValid(bIsValid As Boolean)
'----------------------------------------------------------------------------------------'

    mbIsValid = bIsValid
    RaiseEvent IsValid(mbIsValid)
    
End Property

'----------------------------------------------------------------------------------------'
Public Property Get IsValid() As Boolean
'----------------------------------------------------------------------------------------'

   IsValid = mbIsValid
    
End Property

'----------------------------------------------------------------------------------------'
Private Sub Class_Initialize()
'----------------------------------------------------------------------------------------'

    Set moBrokenRules = New clsBrokenRules
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Class_Terminate()
'----------------------------------------------------------------------------------------'

    Set moBrokenRules = Nothing

End Sub

'----------------------------------------------------------------------------------------'
Private Sub moBrokenRules_BrokenRule()
'----------------------------------------------------------------------------------------'

     IsValid = False
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub moBrokenRules_NoBrokenRules()
'----------------------------------------------------------------------------------------'

     IsValid = True

End Sub

'----------------------------------------------------------------------------------------'
Private Sub RRServerHasAddress()
'----------------------------------------------------------------------------------------'
' http address blank and remote server selected rule validation
'----------------------------------------------------------------------------------------'
    If (mnServerType = eRRServerType.RRRemote) _
            And (msHTTPAddress = "" Or Len(msHTTPAddress) > 255) Then
        'raise this event to inform the display that it is invalid
        moBrokenRules.RuleBroken "HTTPAddress", True
        RaiseEvent ValidHTTPAddress(False)
    Else
        moBrokenRules.RuleBroken "HTTPAddress", False
        'raise this event to inform the display that it is valid
        RaiseEvent ValidHTTPAddress(True)
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Public Property Get ConnectionDetailsChange() As Boolean
'----------------------------------------------------------------------------------------'
' returns whether the user can change connection details
'( this is only for remote server)
'----------------------------------------------------------------------------------------'

    ConnectionDetailsChange = (mnServerType = eRRServerType.RRRemote)

End Property
