Attribute VB_Name = "modMIMsg"
'----------------------------------------------------------------------------------------'
'   File:       modMIMsg.bas
'   Copyright:  InferMed Ltd. 2001. All Rights Reserved
'   Author:     Toby Aldridge, Nov 2001
'   Purpose:    General routines for MIMessage
'----------------------------------------------------------------------------------------'
' REVISIONS
'   NCJ 14 Oct 02 - Changed enumeration names; added Queried and Cancelled for SDVs
'   NCJ 20 Nov 02 - Added (then removed) Freezer statuses
'   TA  20 Jan 03 - GetMIMTypeText now has plural option
'----------------------------------------------------------------------------------------'

Option Explicit

'---------------------------------------------------------------------
Public Function GetStatusText(ByVal nMIMsgType As MIMsgType, ByVal nStatus As Integer) As String
'---------------------------------------------------------------------
' Get the text that represents a MIMsg's status
'---------------------------------------------------------------------

    Select Case nMIMsgType
    Case MIMsgType.mimtDiscrepancy
        Select Case nStatus
        Case eDiscrepancyMIMStatus.dsRaised
            GetStatusText = "Raised"
        Case eDiscrepancyMIMStatus.dsResponded
            GetStatusText = "Responded"
        Case eDiscrepancyMIMStatus.dsClosed
            GetStatusText = "Closed"
        End Select
    Case MIMsgType.mimtSDVMark
        Select Case nStatus
        Case eSDVMIMStatus.ssDone
            GetStatusText = "Done"
        Case eSDVMIMStatus.ssPlanned
            GetStatusText = "Planned"
        Case eSDVMIMStatus.ssCancelled
            GetStatusText = "Cancelled"
        Case eSDVMIMStatus.ssQueried
            GetStatusText = "Queried"
        End Select
    Case MIMsgType.mimtNote
        Select Case nStatus
        Case eNoteMIMStatus.nsPublic
            GetStatusText = "Public"
        Case eNoteMIMStatus.nsPrivate
            GetStatusText = "Private"
        End Select
    End Select
    

End Function

'---------------------------------------------------------------------
Public Function GetMIMTypeText(ByVal nType As MIMsgType, Optional bPlural As Boolean = False) As String
'---------------------------------------------------------------------
' Get the text that represents a MIMessage type
'---------------------------------------------------------------------
Dim sMimType As String
    
    If bPlural Then
        Select Case nType
        Case MIMsgType.mimtDiscrepancy: sMimType = "Discrepancies"
        Case MIMsgType.mimtNote: sMimType = "Notes"
        Case MIMsgType.mimtSDVMark: sMimType = "SDV Marks"
        End Select
    Else
        Select Case nType
        Case MIMsgType.mimtDiscrepancy: sMimType = "Discrepancy"
        Case MIMsgType.mimtNote: sMimType = "Note"
        Case MIMsgType.mimtSDVMark: sMimType = "SDV Mark"
        End Select
    End If
    
    GetMIMTypeText = sMimType

End Function

'---------------------------------------------------------------------
Public Function GetScopeText(ByVal enScope As MIMsgScope) As String
'---------------------------------------------------------------------
' The scope of an MIMessage as a string
'---------------------------------------------------------------------

    Select Case enScope
    Case MIMsgScope.mimscEForm
        GetScopeText = "eForm"
    Case MIMsgScope.mimscQuestion
        GetScopeText = "Question"
    Case MIMsgScope.mimscVisit
        GetScopeText = "Visit"
    Case MIMsgScope.mimscSubject
        GetScopeText = "Subject"
    Case MIMsgScope.mimscStudy
        GetScopeText = "Study"
    End Select

End Function


