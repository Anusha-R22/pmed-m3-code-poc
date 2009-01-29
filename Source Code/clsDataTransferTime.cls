VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataTransferTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Copyright:  InferMed Ltd. 2002. All Rights Reserved
'   File:       clsDataTransferTime.cls
'   Author:     David Hook, April 2002
'   Purpose:    Control the times displayed during data transfer
'--------------------------------------------------------------------------------
' REVISIONS
' DPH 10/05/2002 - Corrected Laboratory Upload Timing
' REM 12/12/02 - added system message timing
' NCJ 19 Dec 02 - Added LFMessages timing
' DPH 27/01/2003 - Added Report Messages Timing
' DPH 24/01/2005 - Added PDU Messages timing (PDU2300)
'--------------------------------------------------------------------------------

Option Explicit
' Running Total
Public LockFreeze As Long
Public StudyUpdate As Long
Public StudyStatus As Long
Public LaboratoryDown As Long
Public MIMessagesDown As Long
Public MIMessagesUp As Long
Public Subject As Long
Public LaboratoryUp As Long
Public SystemMessagesDownload As Long
Public SystemMessagesUpload As Long

Public LFMessagesDown As Long
Public LFMessagesUp As Long

Public ReportMessages As Long

' Seconds allocated to each type of transfer
Public LockFreezeSeconds As Long
Public StudyUpdateSeconds As Long
Public StudyStatusSeconds As Long
Public LaboratoryDownSeconds As Long
Public MIMessagesDownSeconds As Long
Public MIMessagesUpSeconds As Long
Public SubjectSeconds As Long
Public LaboratoryUpSeconds As Long
Public ContingencySeconds As Long
Public SystemMessagesDownSeconds As Long
Public SystemMessagesUpSeconds As Long

Public LFMessagesDownSeconds As Long
Public LFMessagesUpSeconds As Long

Public ReportMessagesSeconds As Long

Public PDUMessagesSeconds As Long

' Total messages to transfer
Public LockFreezeTotal As Long
Public StudyUpdateTotal As Long
Public StudyStatusTotal As Long
Public LaboratoryDownTotal As Long
Public MIMessagesDownTotal As Long
Public MIMessagesUpTotal As Long
Public SubjectTotal As Long
Public LaboratoryUpTotal As Long
Public SystemMessagesDownTotal As Long
Public SystemMessagesUpTotal As Long

Public LFMessagesDownTotal As Long
Public LFMessagesUpTotal As Long

Public ReportMessagesTotal As Long

' Timers
Private mdblLockFreezeTimer As Double
Private mdblStudyUpdateTimer As Double
Private mdblStudyStatusTimer As Double
Private mdblLaboratoryDownTimer As Double
Private mdblMIMessagesDownTimer As Double
Private mdblMIMessagesUpTimer As Double
Private mdblSubjectTimer As Double
Private mdblLaboratoryUpTimer As Double
Private mdblSystemMessagesDownTimer As Double
Private mdblSystemMessagesUpTimer As Double

Private mdblLFMessagesDownTimer As Double
Private mdblLFMessagesUpTimer As Double

Private mdblReportMessagesTimer As Long

Private mdblPDUMessagesTimer As Double

' NCJ 19 Dec 02 - Added LFMessages
Public Enum SectionType
    LockFreezeSection = 1
    StudyUpdateSection = 2
    LaboratoryDownSection = 3
    MIMessagesDownSection = 4
    MIMessagesUpSection = 5
    SubjectSection = 6
    LaboratoryUpSection = 7
    StudyStatusSection = 8
    SystemMessagesDown = 9
    SystemMessagesUp = 10
    LFMessagesDownSection = 11
    LFMessagesUpSection = 12
    ReportFilesSection = 14
    PDUFilesSection = 15
End Enum


'----------------------------------------------------------------------------------------'
Public Sub InitTimes(lLockFreezeSeconds As Long, lStudyUpdateSeconds As Long, _
                    lLaboratoryDownSeconds As Long, lMIMessagesDownSeconds As Long, _
                    lMIMessagesUpSeconds As Long, lSubjectSeconds As Long, _
                    lLaboratoryUpSeconds As Long, lContingencySeconds As Long, _
                    lStudyStatus As Long, lSysMessagesDownload As Long, lSysMessagesUpload As Long)
'----------------------------------------------------------------------------------------'
' Initialise times in seconds
'----------------------------------------------------------------------------------------'

    LockFreezeSeconds = lLockFreezeSeconds
    StudyUpdateSeconds = lStudyUpdateSeconds
    LaboratoryDownSeconds = lLaboratoryDownSeconds
    MIMessagesDownSeconds = lMIMessagesDownSeconds
    MIMessagesUpSeconds = lMIMessagesUpSeconds
    SubjectSeconds = lSubjectSeconds
    LaboratoryUpSeconds = lLaboratoryUpSeconds
    ContingencySeconds = lContingencySeconds
    StudyStatus = lStudyStatus
    SystemMessagesDownSeconds = lSysMessagesDownload
    SystemMessagesUpSeconds = lSysMessagesUpload

End Sub

'----------------------------------------------------------------------------------------'
Public Sub Init(lLockFreezeTotal As Long, lStudyUpdateTotal As Long, _
                lLaboratoryDownTotal As Long, lMIMessagesDownTotal As Long, _
                lMIMessagesUpTotal As Long, lSubjectTotal As Long, _
                lLaboratoryUpTotal As Long, lStudyStatusTotal As Long, _
                lSysMessagesDownTotal As Long, lSysMessagesUpTotal As Long, _
                lLFMessagesDownTotal As Long, lLFMessagesUpTotal As Long, _
                lReportMessagesTotal As Long)
'----------------------------------------------------------------------------------------'
' initialise totals
' NCJ 20/12/02 - added two new parameters for Lock/Freeze messages
'----------------------------------------------------------------------------------------'

    ' set totals
    LockFreezeTotal = lLockFreezeTotal
    StudyUpdateTotal = lStudyUpdateTotal
    LaboratoryDownTotal = lLaboratoryDownTotal
    MIMessagesDownTotal = lMIMessagesDownTotal
    MIMessagesUpTotal = lMIMessagesUpTotal
    SubjectTotal = lSubjectTotal
    LaboratoryUpTotal = lLaboratoryUpTotal
    StudyStatusTotal = lStudyStatusTotal
    SystemMessagesDownTotal = lSysMessagesDownTotal
    SystemMessagesUpTotal = lSysMessagesUpTotal
    LFMessagesDownTotal = lLFMessagesDownTotal
    LFMessagesUpTotal = lLFMessagesUpTotal
    ReportMessagesTotal = lReportMessagesTotal
    
    ' Initialise running totals
    LockFreeze = 0
    StudyUpdate = 0
    LaboratoryDown = 0
    MIMessagesDown = 0
    MIMessagesUp = 0
    Subject = 0
    LaboratoryUp = 0
    StudyStatus = 0
    SystemMessagesDownload = 0
    SystemMessagesUpload = 0
    LFMessagesDown = 0
    LFMessagesUp = 0
    ReportMessages = 0
    
    If Not GetRegSettings Then
        ' default timing settings (based on 56K)
        LockFreezeSeconds = 2
        StudyUpdateSeconds = 60
        LaboratoryDownSeconds = 3
        MIMessagesDownSeconds = 2
        MIMessagesUpSeconds = 2
        SubjectSeconds = 15
        LaboratoryUpSeconds = 2
        ContingencySeconds = 3
        StudyStatusSeconds = 2
        SystemMessagesDownSeconds = 1
        SystemMessagesUpSeconds = 1
        ' NCJ 19 Dec 02 - Made up some numbers here
        LFMessagesDownSeconds = 2
        LFMessagesUpSeconds = 2
        ReportMessagesSeconds = 3
        PDUMessagesSeconds = 0
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub IncrementLockFreeze()
'----------------------------------------------------------------------------------------'
' Increment LockFreeze
'----------------------------------------------------------------------------------------'
    
    LockFreeze = LockFreeze + 1
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub IncrementStudyUpdate()
'----------------------------------------------------------------------------------------'
' Increment StudyUpdate
'----------------------------------------------------------------------------------------'

    StudyUpdate = StudyUpdate + 1
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub IncrementStudyStatus()
'----------------------------------------------------------------------------------------'
' Increment StudyStatus
'----------------------------------------------------------------------------------------'

    StudyStatus = StudyStatus + 1
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub IncrementLaboratoryDown()
'----------------------------------------------------------------------------------------'
' Increment LaboratoryDown
'----------------------------------------------------------------------------------------'

    LaboratoryDown = LaboratoryDown + 1
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub IncrementSystemMessagesDown()
'----------------------------------------------------------------------------------------'
'Increment SystemMessagesDownload
'----------------------------------------------------------------------------------------'
    SystemMessagesDownload = SystemMessagesDownload + 1
End Sub

'----------------------------------------------------------------------------------------'
Public Sub IncrementSystemMessagesUp()
'----------------------------------------------------------------------------------------'
'Increment SystemMessagesUpload
'----------------------------------------------------------------------------------------'
    SystemMessagesUpload = SystemMessagesUpload + 1
End Sub

'----------------------------------------------------------------------------------------'
Public Sub IncrementMIMessagesDown()
'----------------------------------------------------------------------------------------'
' Increment MIMessagesDown
'----------------------------------------------------------------------------------------'

    MIMessagesDown = MIMessagesDown + 1
    
    ' stop timing if completed
    If MIMessagesDown = MIMessagesDownTotal Then
        EndMIMessagesDown
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub IncrementMIMessagesUp()
'----------------------------------------------------------------------------------------'
' Increment MIMessagesUp
'----------------------------------------------------------------------------------------'

    MIMessagesUp = MIMessagesUp + 1
 
    ' stop timing if completed
    If MIMessagesUp = MIMessagesUpTotal Then
        EndMIMessagesUp
    End If
 
End Sub

'----------------------------------------------------------------------------------------'
Public Sub IncrementLFMessagesDown()
'----------------------------------------------------------------------------------------'
' Increment MIMessagesDown
'----------------------------------------------------------------------------------------'

    LFMessagesDown = LFMessagesDown + 1
    
    ' stop timing if completed
    If LFMessagesDown = LFMessagesDownTotal Then
        EndLFMessagesDown
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub IncrementLFMessagesUp()
'----------------------------------------------------------------------------------------'
' Increment MIMessagesUp
'----------------------------------------------------------------------------------------'

    LFMessagesUp = LFMessagesUp + 1
 
    ' stop timing if completed
    If LFMessagesUp = LFMessagesUpTotal Then
        EndLFMessagesUp
    End If
 
End Sub


'----------------------------------------------------------------------------------------'
Public Sub IncrementSubject()
'----------------------------------------------------------------------------------------'
' Increment Subject
'----------------------------------------------------------------------------------------'

    Subject = Subject + 1
    
    ' stop timing if completed
    If Subject = SubjectTotal Then
        EndSubjects
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Public Sub IncrementLaboratoryUp()
'----------------------------------------------------------------------------------------'
' Increment LaboratoryUp
'----------------------------------------------------------------------------------------'

    LaboratoryUp = LaboratoryUp + 1
    
    ' stop timing if completed
    If LaboratoryUp = LaboratoryUpTotal Then
        EndLaboratoryUp
    End If

End Sub

'----------------------------------------------------------------------------------------'
Public Property Get RemainingTime() As Long
'----------------------------------------------------------------------------------------'
' Calculate remaining time (in seconds)
' NCJ 19 Dec 02 - Added LFMessages
'----------------------------------------------------------------------------------------'
Dim lRemainingTime As Long

    lRemainingTime = ((LockFreezeTotal - LockFreeze) * LockFreezeSeconds) + _
                     ((StudyUpdateTotal - StudyUpdate) * StudyUpdateSeconds) + _
                     ((StudyStatusTotal - StudyStatus) * StudyStatusSeconds) + _
                     ((LaboratoryDownTotal - LaboratoryDown) * LaboratoryDownSeconds) + _
                     ((MIMessagesDownTotal - MIMessagesDown) * MIMessagesDownSeconds) + _
                     ((MIMessagesUpTotal - MIMessagesUp) * MIMessagesUpSeconds) + _
                     ((SubjectTotal - Subject) * SubjectSeconds) + _
                     ((LaboratoryUpTotal - LaboratoryUp) * LaboratoryUpSeconds) + _
                     ((SystemMessagesDownTotal - SystemMessagesDownload) * SystemMessagesDownSeconds) + _
                     ((SystemMessagesUpTotal - SystemMessagesUpload) * SystemMessagesUpSeconds) + _
                     ((LFMessagesDownTotal - LFMessagesDown) * LFMessagesDownSeconds) + _
                     ((LFMessagesUpTotal - LFMessagesUp) * LFMessagesUpSeconds) + _
                     (PDUMessagesSeconds) + _
                     ContingencySeconds
    
    RemainingTime = lRemainingTime
    
End Property

'----------------------------------------------------------------------------------------'
Private Sub StartLockFreeze()
'----------------------------------------------------------------------------------------'
' start timing
'----------------------------------------------------------------------------------------'
    mdblLockFreezeTimer = Timer
End Sub

'----------------------------------------------------------------------------------------'
Private Sub EndLockFreeze()
'----------------------------------------------------------------------------------------'
' end timing & set seconds taken
'----------------------------------------------------------------------------------------'
    LockFreezeSeconds = SetMinAtLeastOne(CLng(Timer - mdblLockFreezeTimer))
End Sub

'----------------------------------------------------------------------------------------'
Private Sub StartStudyUpdate()
'----------------------------------------------------------------------------------------'
' start timing
'----------------------------------------------------------------------------------------'
    mdblStudyUpdateTimer = Timer
End Sub

'----------------------------------------------------------------------------------------'
Private Sub EndStudyUpdate()
'----------------------------------------------------------------------------------------'
' end timing & set seconds taken
'----------------------------------------------------------------------------------------'
    StudyUpdateSeconds = SetMinAtLeastOne(CLng(Timer - mdblStudyUpdateTimer))
End Sub

'----------------------------------------------------------------------------------------'
Private Sub StartStudyStatus()
'----------------------------------------------------------------------------------------'
' start timing
'----------------------------------------------------------------------------------------'
    mdblStudyStatusTimer = Timer
End Sub

'----------------------------------------------------------------------------------------'
Private Sub EndStudyStatus()
'----------------------------------------------------------------------------------------'
' end timing & set seconds taken
'----------------------------------------------------------------------------------------'
    StudyStatusSeconds = SetMinAtLeastOne(CLng(Timer - mdblStudyStatusTimer))
End Sub

'----------------------------------------------------------------------------------------'
Private Sub StartLaboratoryDown()
'----------------------------------------------------------------------------------------'
' start timing
'----------------------------------------------------------------------------------------'
    mdblLaboratoryDownTimer = Timer
End Sub

'----------------------------------------------------------------------------------------'
Private Sub EndLaboratoryDown()
'----------------------------------------------------------------------------------------'
' end timing & set seconds taken
'----------------------------------------------------------------------------------------'
    LaboratoryDownSeconds = SetMinAtLeastOne(CLng(Timer - mdblLaboratoryDownTimer))
End Sub

'----------------------------------------------------------------------------------------'
Private Sub StartMIMessagesDown()
'----------------------------------------------------------------------------------------'
' start timing
'----------------------------------------------------------------------------------------'
    mdblMIMessagesDownTimer = Timer
End Sub

'----------------------------------------------------------------------------------------'
Private Sub EndMIMessagesDown()
'----------------------------------------------------------------------------------------'
' end timing & set seconds taken
'----------------------------------------------------------------------------------------'
    MIMessagesDownSeconds = SetMinAtLeastOne(CLng((Timer - mdblMIMessagesDownTimer) / MIMessagesDownTotal))
End Sub

'----------------------------------------------------------------------------------------'
Private Sub StartMIMessagesUp()
'----------------------------------------------------------------------------------------'
' start timing
'----------------------------------------------------------------------------------------'
    mdblMIMessagesUpTimer = Timer
End Sub

'----------------------------------------------------------------------------------------'
Private Sub EndMIMessagesUp()
'----------------------------------------------------------------------------------------'
' end timing & set seconds taken
'----------------------------------------------------------------------------------------'
    MIMessagesUpSeconds = SetMinAtLeastOne(CLng((Timer - mdblMIMessagesUpTimer) / MIMessagesUpTotal))
End Sub

'----------------------------------------------------------------------------------------'
Private Sub StartLFMessagesDown()
'----------------------------------------------------------------------------------------'
' start timing
'----------------------------------------------------------------------------------------'
    mdblLFMessagesDownTimer = Timer
End Sub

'----------------------------------------------------------------------------------------'
Private Sub EndLFMessagesDown()
'----------------------------------------------------------------------------------------'
' end timing & set seconds taken
'----------------------------------------------------------------------------------------'
    LFMessagesDownSeconds = SetMinAtLeastOne(CLng((Timer - mdblLFMessagesDownTimer) / LFMessagesDownTotal))
End Sub

'----------------------------------------------------------------------------------------'
Private Sub StartLFMessagesUp()
'----------------------------------------------------------------------------------------'
' start timing
'----------------------------------------------------------------------------------------'
    mdblLFMessagesUpTimer = Timer
End Sub

'----------------------------------------------------------------------------------------'
Private Sub EndLFMessagesUp()
'----------------------------------------------------------------------------------------'
' end timing & set seconds taken
'----------------------------------------------------------------------------------------'
    LFMessagesUpSeconds = SetMinAtLeastOne(CLng((Timer - mdblLFMessagesUpTimer) / LFMessagesUpTotal))
End Sub

'----------------------------------------------------------------------------------------'
Private Sub StartReportMessages()
'----------------------------------------------------------------------------------------'
' start timing
'----------------------------------------------------------------------------------------'
    mdblReportMessagesTimer = Timer
End Sub

'----------------------------------------------------------------------------------------'
Private Sub EndReportMessages()
'----------------------------------------------------------------------------------------'
' end timing & set seconds taken
'----------------------------------------------------------------------------------------'
    ReportMessagesSeconds = SetMinAtLeastOne(CLng(Timer - mdblReportMessagesTimer))
End Sub

'----------------------------------------------------------------------------------------'
Private Sub StartSubjects()
'----------------------------------------------------------------------------------------'
' start timing
'----------------------------------------------------------------------------------------'
    mdblSubjectTimer = Timer
End Sub

'----------------------------------------------------------------------------------------'
Private Sub EndSubjects()
'----------------------------------------------------------------------------------------'
' end timing & set seconds taken
'----------------------------------------------------------------------------------------'
    SubjectSeconds = SetMinAtLeastOne(CLng((Timer - mdblSubjectTimer) / SubjectTotal))
End Sub

'----------------------------------------------------------------------------------------'
Private Sub StartLaboratoryUp()
'----------------------------------------------------------------------------------------'
' start timing
'----------------------------------------------------------------------------------------'
    mdblLaboratoryUpTimer = Timer
End Sub

'----------------------------------------------------------------------------------------'
Private Sub EndLaboratoryUp()
'----------------------------------------------------------------------------------------'
' end timing & set seconds taken
'----------------------------------------------------------------------------------------'
' REVISIONS
' DPH 10/05/2002 - Corrected Laboratory Upload Timing
'----------------------------------------------------------------------------------------'

    LaboratoryUpSeconds = SetMinAtLeastOne(CLng((Timer - mdblLaboratoryUpTimer) / LaboratoryUpTotal))
End Sub

'----------------------------------------------------------------------------------------'
Private Sub StartSystemMessageDown()
'----------------------------------------------------------------------------------------'
'Start timing
'----------------------------------------------------------------------------------------'
    mdblSystemMessagesDownTimer = Timer
End Sub

'----------------------------------------------------------------------------------------'
Private Sub EndSystemMessageDown()
'----------------------------------------------------------------------------------------'
' end timing & set seconds taken
'----------------------------------------------------------------------------------------'

    SystemMessagesDownSeconds = SetMinAtLeastOne(CLng((Timer - mdblSystemMessagesDownTimer) / SystemMessagesDownTotal))
End Sub

'----------------------------------------------------------------------------------------'
Private Sub StartSystemMessageUp()
'----------------------------------------------------------------------------------------'
'Start timing
'----------------------------------------------------------------------------------------'
    mdblSystemMessagesUpTimer = Timer
End Sub

'----------------------------------------------------------------------------------------'
Private Sub EndSystemMessageUp()
'----------------------------------------------------------------------------------------'
' end timing & set seconds taken
'----------------------------------------------------------------------------------------'

    SystemMessagesUpSeconds = SetMinAtLeastOne(CLng((Timer - mdblSystemMessagesUpTimer) / SetMinAtLeastOne(SystemMessagesUpTotal)))
End Sub

'----------------------------------------------------------------------------------------'
Private Sub StartPDUMessages()
'----------------------------------------------------------------------------------------'
' start timing
'----------------------------------------------------------------------------------------'
    mdblPDUMessagesTimer = Timer
End Sub

'----------------------------------------------------------------------------------------'
Private Sub EndPDUMessages()
'----------------------------------------------------------------------------------------'
' end timing & set seconds taken
'----------------------------------------------------------------------------------------'
    PDUMessagesSeconds = SetMinAtLeastOne(CLng(Timer - mdblPDUMessagesTimer))
End Sub

'----------------------------------------------------------------------------------------'
Private Function SetMinAtLeastOne(lValue As Long) As Long
'----------------------------------------------------------------------------------------'
' set min as one
'----------------------------------------------------------------------------------------'
    If lValue < 1 Then
        SetMinAtLeastOne = 1
    Else
        SetMinAtLeastOne = lValue
    End If
End Function

'----------------------------------------------------------------------------------------'
Public Sub StartTiming(eSection As SectionType)
'----------------------------------------------------------------------------------------'
' Start timing of chosen section
' NCJ 19 Dec 02 - Added LFMessages
'----------------------------------------------------------------------------------------'
    
    Select Case eSection
        Case SectionType.LockFreezeSection
            StartLockFreeze
        Case SectionType.LaboratoryDownSection
            StartLaboratoryDown
        Case SectionType.StudyUpdateSection
            StartStudyUpdate
        Case SectionType.MIMessagesDownSection
            StartMIMessagesDown
        Case SectionType.SubjectSection
            StartSubjects
        Case SectionType.MIMessagesUpSection
            StartMIMessagesUp
        Case SectionType.LaboratoryUpSection
            StartLaboratoryUp
        Case SectionType.SystemMessagesDown
            StartSystemMessageDown
        Case SectionType.SystemMessagesUp
            StartSystemMessageUp
        Case SectionType.LFMessagesDownSection
            StartLFMessagesDown
        Case SectionType.LFMessagesUpSection
            StartLFMessagesUp
        Case SectionType.ReportFilesSection
            StartReportMessages
        Case SectionType.PDUFilesSection
            StartPDUMessages
    End Select

End Sub

'----------------------------------------------------------------------------------------'
Public Sub StopTiming(eSection As SectionType)
'----------------------------------------------------------------------------------------'
' Stop timing of chosen section
' NCJ 19 Dec 02 - Added LFMessages
'----------------------------------------------------------------------------------------'
    Select Case eSection
        Case SectionType.LockFreezeSection
            EndLockFreeze
        Case SectionType.LaboratoryDownSection
            EndLaboratoryDown
        Case SectionType.StudyUpdateSection
            EndStudyUpdate
        Case SectionType.MIMessagesDownSection
            EndMIMessagesDown
        Case SectionType.SubjectSection
            EndSubjects
        Case SectionType.MIMessagesUpSection
            EndMIMessagesUp
        Case SectionType.LaboratoryUpSection
            EndLaboratoryUp
        Case SectionType.SystemMessagesDown
            EndSystemMessageDown
        Case SectionType.SystemMessagesUp
            EndSystemMessageUp
        Case SectionType.LFMessagesDownSection
            EndLFMessagesDown
        Case SectionType.LFMessagesUpSection
            EndLFMessagesUp
        Case SectionType.ReportFilesSection
            EndReportMessages
        Case SectionType.PDUFilesSection
            EndPDUMessages
    End Select
End Sub

'----------------------------------------------------------------------------------------'
Private Function GetRegSettings() As Boolean
'----------------------------------------------------------------------------------------'
' Pick up previous settings from Registry (if exist)
'REM 12/12/02 - added system message settings
'----------------------------------------------------------------------------------------'
Dim sSetting As String
Dim vSetting As Variant

    'jump out if error
    On Error GoTo ErrHandler
    
    sSetting = GetSetting(GetApplicationTitle, "Data Transfer", "Timings", "")

    If sSetting = "" Then
        'not found
        GetRegSettings = False
    Else
        vSetting = Split(sSetting, ",")
        ' Retrieve previously saved settings
        LockFreezeSeconds = vSetting(0)
        StudyUpdateSeconds = vSetting(1)
        LaboratoryDownSeconds = vSetting(2)
        MIMessagesDownSeconds = vSetting(3)
        MIMessagesUpSeconds = vSetting(4)
        SubjectSeconds = vSetting(5)
        LaboratoryUpSeconds = vSetting(6)
        ContingencySeconds = vSetting(7)
        StudyStatusSeconds = vSetting(8)
        'check to make sure there are system message settings
        If UBound(vSetting) > 8 Then
            SystemMessagesDownSeconds = vSetting(9)
            SystemMessagesUpSeconds = vSetting(10)
        Else 'set defaults
            SystemMessagesDownSeconds = 2
            SystemMessagesUpSeconds = 2
        End If
         'check to make sure there are system message settings
        If UBound(vSetting) > 11 Then
            LFMessagesDownSeconds = vSetting(11)
            LFMessagesUpSeconds = vSetting(12)
        Else
            LFMessagesDownSeconds = 2
            LFMessagesUpSeconds = 2
        End If
        If UBound(vSetting) > 12 Then
            ReportMessagesSeconds = vSetting(13)
        Else
            ReportMessagesSeconds = 3
        End If
        If UBound(vSetting) > 13 Then
            PDUMessagesSeconds = vSetting(14)
        Else
            ' store zero as not all people use pdu and
            ' so if subsequent transfers occur will use stored setting
            PDUMessagesSeconds = 0
        End If
        GetRegSettings = True
    End If
    
Exit Function
ErrHandler:
    GetRegSettings = False
End Function

'----------------------------------------------------------------------------------------'
Public Sub SaveRegSettings()
'----------------------------------------------------------------------------------------'
' save new seconds settings to Registry
'----------------------------------------------------------------------------------------'
On Error GoTo ErrHandler
Dim sSetting As String

        ' Put together setting and save to registry for future transfers
        sSetting = LockFreezeSeconds & "," & StudyUpdateSeconds & "," & LaboratoryDownSeconds & "," & _
                MIMessagesDownSeconds & "," & MIMessagesUpSeconds & "," & SubjectSeconds & "," & _
                LaboratoryUpSeconds & "," & ContingencySeconds & "," & StudyStatusSeconds & "," & _
                SystemMessagesDownSeconds & "," & SystemMessagesUpSeconds & "," & _
                LFMessagesDownSeconds & "," & LFMessagesUpSeconds & "," & _
                ReportMessagesSeconds & "," & PDUMessagesSeconds
        
        Call SaveSetting(GetApplicationTitle, "Data Transfer", "Timings", sSetting)

Exit Sub
ErrHandler:

End Sub
