Attribute VB_Name = "modMIMessages"
'----------------------------------------------------------------------------------------'
'   File:       modMIMessages.bas
'   Copyright:  InferMed Ltd. 2000. All Rights Reserved
'   Author:     Nicky Johns, May 2000
'   Purpose:    Routines for Bidirectional Communication
'               using the MIMessage table
'----------------------------------------------------------------------------------------'
' Revisions:
'   NCJ 17-22 May 2000 - Initial development
'   NCJ 21/11/00 - Do not use module-level MIMessage objects
'   NCJ 24/11/00 SR 4021 - Do RemoveNull when reading ResponseValue
'   TA 26/08/2002: REmoved all old code that has been replaced by the business object
'   TA 26/08/2002: Added support for MIMsg statuses
'   TA 28/08/2002: Added support for Note statuses
'   RS 01/10/2002: Added Timezone support
'   NCJ 14 Oct 02 - Allow different scopes for SDVs (with other message details being optional)
'   NCJ 5 Nov 02 - New SDVExists function
'   ic 14/07/2006  issue 2597 because the check to see if an SDV already exists on an item is done
'                   before the item is locked it is possible for 2 users to add an SDV to the same item
'----------------------------------------------------------------------------------------'

Option Explicit

'---------------------------------------------------------------------------------
Public Function CreateNewMIMessage(MType As MIMsgType, _
                                enScope As MIMsgScope, _
                                dblTimeStamp As Double, nTimezoneOffset As Integer, _
                                sSite As String, _
                                lTrialId As Long, _
                                lPersonId As Long, _
                                oResponse As Response, _
                                Optional lVisitId As Long = 0, _
                                Optional nVisitCycle As Integer = 0, _
                                Optional lCRFPageTaskId As Long = 0, _
                                Optional lResponseTaskId As Long = 0, _
                                Optional nResponseCycle As Integer = 0, _
                                Optional sResponseValue As String = "", _
                                Optional dblResponseTimeStamp As Double = 0, _
                                Optional lEFormId As Long = 0, _
                                Optional nEFormCycle As Integer = 0, _
                                Optional lQuestionId As Long = 0, _
                                Optional sDataUserName As String = "", _
                                Optional oSubject As StudySubject = Nothing) As Boolean
'---------------------------------------------------------------------------------
' Handle everything you need to do to create a new discrepancy/SDV Mark/message/note
' NB To respond to a discrepancy or remark, use CreateDiscrepancyResponse
' Ask user for details and create new MIMessage record in MIMessage table
' NCJ 21/11/00 - Find out the ResponseTimeStamp and store it too
' RS 30/09/2002 Added Timestamp & Timezone parameter
' TA 02/10/2002 If oSubject is passed in then its Disc/SDV/Note Status is updated
' NCJ 14 Oct 02 - Added enScope argument (actually only for SDVs); made many arguments optional
' NCJ 16 Oct 02 - Return FALSE if the user cancelled the NewMIMessage dialog
' ic 14/07/2006  issue 2597 because the check to see if an SDV already exists on an item is done
'                before the item is locked it is possible for 2 users to add an SDV to the same item
'---------------------------------------------------------------------------------
Dim sMsgText As String
Dim nPriority As Integer
Dim lOCDiscrepancyID As Long
Dim sObjName As String
Dim sTrialName As String
Dim nStatus As Integer
Dim oDisc As MIDiscrepancy
Dim oSDV As MISDV
Dim oNote As MINote
Dim oLockForMIMsg As clsLockForMIMsg

    On Error GoTo Errlabel
    
    sTrialName = TrialNameFromId(lTrialId)
    
    
    Set oLockForMIMsg = New clsLockForMIMsg
    
    If oLockForMIMsg.LockIfNeeded(sTrialName, sSite, lPersonId, MType, Nothing, oResponse) Then
        
        Select Case enScope
        Case MIMsgScope.mimscSubject
            ' Get subject name
            sObjName = SubjectLabelFromTrialSiteId(lTrialId, sSite, lPersonId)
            If sObjName = "" Then
                ' No label - use the Person ID
                sObjName = lPersonId
            End If
        Case MIMsgScope.mimscVisit
            ' Get visit name (includes Cycle no.)
            sObjName = VisitNameFromId(lTrialId, lVisitId, nVisitCycle)
        Case MIMsgScope.mimscEForm
            ' Get form Name (includes Cycle no.)
            sObjName = CRFPageNameFromTaskId(lTrialId, sSite, lPersonId, lCRFPageTaskId)
        Case MIMsgScope.mimscQuestion
            ' Get data item Name (includes Cycle no.)
            sObjName = DataItemNameFromTaskId(lTrialId, sSite, lPersonId, lResponseTaskId, nResponseCycle)
        End Select
        
        ' Ask the user for their input
        If frmNewMIMessage.Display(MType, enScope, sTrialName, sObjName, sResponseValue, _
                            sMsgText, lOCDiscrepancyID, nPriority, nStatus) Then
    
            Select Case MType
            Case MIMsgType.mimtDiscrepancy
                Set oDisc = New MIDiscrepancy
                oDisc.Raise gsADOConnectString, sMsgText, nPriority, lOCDiscrepancyID, goUser.UserName, _
                            goUser.UserNameFull, GetMIMsgSource, enScope, _
                            sTrialName, sSite, lPersonId, lVisitId, _
                            nVisitCycle, lCRFPageTaskId, lResponseTaskId, nResponseCycle, _
                            dblResponseTimeStamp, sResponseValue, _
                            lEFormId, nEFormCycle, lQuestionId, sDataUserName, _
                            dblTimeStamp, nTimezoneOffset
                oDisc.Save
                
                'update the discrepancy statuses
                Call UpdateMIMsgStatus(gsADOConnectString, MIMsgType.mimtDiscrepancy, _
                            sTrialName, lTrialId, sSite, lPersonId, lVisitId, _
                            nVisitCycle, lCRFPageTaskId, lResponseTaskId, nResponseCycle, oSubject)
                            
            Case MIMsgType.mimtSDVMark
                'ic 14/07/2006 issue 2597 because the check to see if an SDV already exists on an item
                'is done before the item is locked it is possible for 2 users to add an SDV to the same
                'item. this added check after the lock is in place will prevent this
                If SDVExists(enScope, sTrialName, sSite, lPersonId, lVisitId, nVisitCycle, lCRFPageTaskId, lResponseTaskId) Then
                    DialogInformation "An SDV already exists for this item"
                    
                Else
                    Set oSDV = New MISDV
                    oSDV.Raise gsADOConnectString, sMsgText, (nStatus), goUser.UserName, _
                                goUser.UserNameFull, GetMIMsgSource, _
                                dblTimeStamp, nTimezoneOffset, enScope, _
                                sTrialName, sSite, lPersonId, lVisitId, _
                                nVisitCycle, lCRFPageTaskId, lResponseTaskId, nResponseCycle, _
                                lEFormId, nEFormCycle, lQuestionId, sDataUserName, _
                                dblResponseTimeStamp, sResponseValue
                    oSDV.Save
                    
                    'Update MIMsgStatus
                    Call UpdateMIMsgStatus(gsADOConnectString, MIMsgType.mimtSDVMark, _
                                sTrialName, lTrialId, sSite, lPersonId, lVisitId, _
                                nVisitCycle, lCRFPageTaskId, lResponseTaskId, nResponseCycle, oSubject)
                End If
        
            Case MIMsgType.mimtNote
                Set oNote = New MINote
                oNote.Init gsADOConnectString, sMsgText, goUser.UserName, _
                            goUser.UserNameFull, GetMIMsgSource, enScope, _
                            sTrialName, sSite, lPersonId, lVisitId, _
                            nVisitCycle, lCRFPageTaskId, lResponseTaskId, nResponseCycle, _
                            dblResponseTimeStamp, sResponseValue, _
                            lEFormId, nEFormCycle, lQuestionId, sDataUserName, _
                            dblTimeStamp, nTimezoneOffset, (nStatus)
                oNote.Save
                'update the note statuses
                Call UpdateNoteStatus(gsADOConnectString, mimscQuestion, sTrialName, lTrialId, sSite, lPersonId, lVisitId, _
                            nVisitCycle, lCRFPageTaskId, lResponseTaskId, nResponseCycle, oSubject)
            End Select
    
            'update SDV and discrepancy count
            frmMenu.UpdateDiscCount
    
            CreateNewMIMessage = True
        Else
            ' They cancelled
            CreateNewMIMessage = False
        End If
        
        'let's unlock
        Call oLockForMIMsg.UnlockIfNeeded
    Else
        CreateNewMIMessage = False
    End If
    
Exit Function
Errlabel:
    'TA 29/04/2003:
    'let's unlock just in case -does nothing if ther e is no lock
    'we need to store errors as they will be cleared
    Dim lErr As Long
    Dim sErr As String
    lErr = Err.Number
    sErr = Err.Description
    If Not oLockForMIMsg Is Nothing Then
        Call oLockForMIMsg.UnlockIfNeeded
    End If
    
    Err.Raise lErr, , sErr & "|modMIMessages.CreateNewMIMessage"

End Function

'----------------------------------------------------------------------------------------'
Public Function GetMIMsgSource() As MIMsgSource
'----------------------------------------------------------------------------------------'
' Whether we are on a site or at the Server
'----------------------------------------------------------------------------------------'

    If gblnRemoteSite Then
        GetMIMsgSource = mimsSite
    Else
        GetMIMsgSource = mimsServer
    End If

End Function

'----------------------------------------------------------------------------------------'
Public Function SDVExists(nScope As MIMsgScope, _
                        sStudyName As String, sSite As String, lSubjectId As Long, _
                        Optional lVisitId As Long = 0, Optional nVisitCycle As Integer = 0, _
                        Optional leFormTaskId As Long = 0, _
                        Optional lResponseId As Long = 0) As Boolean
'----------------------------------------------------------------------------------------'
' NCJ 5 Nov 02 - Does an SDV already exist for this object?
' Note that the optional parameters should be supplied according to the Scope
'----------------------------------------------------------------------------------------'
Dim oMIMData As MIDataLists
Dim lObjectId As Long
Dim nObjectSource As Integer

    Set oMIMData = New MIDataLists
    
    SDVExists = oMIMData.MIMessageExists(gsADOConnectString, MIMsgType.mimtSDVMark, _
                            nScope, sStudyName, sSite, lSubjectId, _
                            lObjectId, nObjectSource, _
                            lVisitId, nVisitCycle, _
                            leFormTaskId, lResponseId)
    
    Set oMIMData = Nothing

End Function
