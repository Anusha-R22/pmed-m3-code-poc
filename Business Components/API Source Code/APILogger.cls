VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "APILogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private msMessages As String
Private mbUseFile As Boolean



Public Sub Init()


    mbUseFile = False
    msMessages = ""
    On Error GoTo ErrLabel
    StringFromFile App.Path & "\API.log"
    mbUseFile = True
    
    
    Exit Sub
    
    
ErrLabel:
    mbUseFile = False
    

End Sub



Public Sub Log(sMessage As String)

    msMessages = msMessages & vbCrLf & Now & " :" & vbTab & sMessage
    If mbUseFile Then
        On Error Resume Next
        StringToFileAppend App.Path & "\API.log", Now & " :" & vbTab & sMessage
    End If
End Sub

Public Property Get Messages() As String
    Messages = msMessages
End Property

'----------------------------------------------------------------------------------------'
Private Sub StringToFileAppend(sFileName As String, sText As String)
'----------------------------------------------------------------------------------------'
' Write string to given file
'----------------------------------------------------------------------------------------'
Dim n As Integer

    n = FreeFile
    Open sFileName For Append As n
    
    Print #n, sText
    
    Close n

End Sub
