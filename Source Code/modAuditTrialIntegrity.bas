Attribute VB_Name = "modAuditTrialIntegrity"
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999. All Rights Reserved
'   File:       AuditTrialIntegrity.bas
'   Author:     Stephen Morris
'   Purpose:
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   1   17/12/99    SDM Creation
'----------------------------------------------------------------------------------------'
Option Explicit
Private Const msAuditTrailFile = "AuTrIntg.txt"

'----------------------------------------------------------------------------------------'
Public Function CheckDataItemResponseHistory() As String
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'
Dim rsStudy As ADODB.Recordset
Dim rsTrialSite As ADODB.Recordset
Dim rsStatus As ADODB.Recordset
Dim sSQL As String
Dim nChangedLoop As Integer
Dim nFileNumber As Integer
Dim sChangedText As String
    
    nFileNumber = FreeFile
    'changed Mo Morris 10/2/00
    Open gsTEMP_PATH & msAuditTrailFile For Output As #nFileNumber
    'Open App.Path & "\AuTrIntg.txt" For Output As #nFileNumber
    
    'Retrive all Study Ids
    sSQL = "SELECT " & _
           "DISTINCT(ClinicalTrialId) " & _
           "FROM " & _
           "DataItemResponseHistory"
    Set rsStudy = New ADODB.Recordset
    rsStudy.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
    
    'Loop each Study Id
    'SDM 01/02/00   Added the <And Not IsEmpty(rsStudy.Fields("ClinicalTrialId").Value)>
    'REM 20/02/03 - Changed from one If statment to two
    If Not rsStudy.EOF Then
        If Not IsEmpty(rsStudy.Fields("ClinicalTrialId").Value) Then
            rsStudy.MoveFirst
            Do While Not rsStudy.EOF
            
                'Retrieve all TrialSite for a Study Id
                sSQL = "SELECT " & _
                       "DISTINCT TrialSite " & _
                       "FROM " & _
                       "DataItemResponseHistory " & _
                       "WHERE " & _
                       "ClinicalTrialId = " & rsStudy.Fields("ClinicalTrialId").Value
                Set rsTrialSite = New ADODB.Recordset
                rsTrialSite.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
                
                'Loop each TrialSite
                If Not rsTrialSite.EOF Then
                    rsTrialSite.MoveFirst
                    Do While Not rsTrialSite.EOF
                    
                        'Loop each possible Changed value
                        For nChangedLoop = 0 To 2
                        
                            'Get display text for changed type
                            Select Case nChangedLoop
                                Case Changed.Changed
                                    sChangedText = "Data changed but not sent"
                                Case Changed.Imported
                                    sChangedText = "Data received"
                                Case Changed.NoChange
                                    sChangedText = "Data Sent"
                                Case Else
                                    sChangedText = "Unknown"
                            End Select
                            
                            'Get record count for given Study Id, TrialSite and Changed value
                            sSQL = "SELECT " & _
                                   "COUNT(*) AS Total " & _
                                   "FROM " & _
                                   "DataItemResponseHistory " & _
                                   "WHERE " & _
                                   "ClinicalTrialId = " & rsStudy.Fields("ClinicalTrialId").Value & " AND " & _
                                   "TrialSite = '" & rsTrialSite.Fields("TrialSite").Value & "' AND " & _
                                   "Changed = " & nChangedLoop
                            Set rsStatus = New ADODB.Recordset
                            rsStatus.Open sSQL, MacroADODBConnection, adOpenKeyset, adLockReadOnly, adCmdText
                            
                            'Output information to text file
                            If Not rsStatus.EOF Then
                                rsStatus.MoveFirst
                                Do While Not rsStatus.EOF
                                    Print #nFileNumber, "Study = " & rsStudy.Fields("ClinicalTrialId").Value & ", " & _
                                                        "TrialSite = " & rsTrialSite.Fields("TrialSite").Value & ", " & _
                                                        "Status = " & sChangedText & ", " & _
                                                        "No. of Records = " & rsStatus.Fields("Total").Value
                                    rsStatus.MoveNext
                                Loop
                            End If
                                   
                        Next nChangedLoop
            
                        rsTrialSite.MoveNext
                    Loop
                End If
            
                rsStudy.MoveNext
            Loop
        End If
    End If
    
    Close #nFileNumber
    
    'Path returned to be used for launching Notepad
    'changed Mo Morris 10/2/00
    CheckDataItemResponseHistory = gsTEMP_PATH & msAuditTrailFile
    'CheckDataItemResponseHistory = App.Path & "\AuTrIntg.txt"
End Function

'----------------------------------------------------------------------------------------'
Public Sub ResetTransferFlags()
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'
Dim sSQL As String
Dim vTableName As Variant
    'Reset Changed field in listed tables
    For Each vTableName In Array("TrialSubject", _
                                 "VisitInstance", _
                                 "CRFPageInstance", _
                                 "DataItemResponse", _
                                 "DataItemResponseHistory")
        sSQL = "UPDATE " & _
               vTableName & " " & _
               "SET Changed = " & Changed.Changed
        MacroADODBConnection.Execute sSQL
    Next vTableName
    MsgBox "All transfer statuses have been reset to Changed"
End Sub
