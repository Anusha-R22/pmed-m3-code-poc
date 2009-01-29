Attribute VB_Name = "modArezzo"
'----------------------------------------------------------------------------------------'
'   File:       modArezzo.bas
'   Copyright:  InferMed Ltd. 1999-2006. All Rights Reserved
'   Author:     Nicky Johns, September 1999
'   Purpose:    Arezzo Engine & Data routines for MACRO Data Management
'----------------------------------------------------------------------------------------'
'----------------------------------------------------------------------------------------'
' Revisions:
'
' NCJ 25 Sep 01 - Module re-vamped for MACRO 2.2
'       Much code in Public routines removed; private routines deleted
'       Irrelevant revision comments removed
' NCJ 1 Oct 01 - Date parsing code from ProformaEngine moved to here
'       Added handlers for non-MACRO Arezzo tasks
' NCJ 22 Mar 06 - Added code to set "MACRO settings" in AREZZO
'----------------------------------------------------------------------------------------'

Option Explicit

' What's returned from EvaluateExpression if it fails
Public Const gArezzoSyntaxError = "Syntax error"

'----------------------------------------------------------------
Public Function SetAREZZOSetting(sKey As String, sValue As String)
'----------------------------------------------------------------
' NCJ 22 Mar 06 - Assert a macro_setting/2 clause for AREZZO's use
'----------------------------------------------------------------
Dim sGoal As String
Dim sR As String
Dim sMacroSetting As String

    sMacroSetting = "macro_setting( " & LCase(sKey) & ", "
    
    sGoal = "retractall( " & sMacroSetting & "_ )), "
    sGoal = sGoal & "assert( " & sMacroSetting & LCase(sValue) & " )), "
    sGoal = sGoal & "write('0000'). "
    
    Call goArezzo.ALM.GetPrologResult(sGoal, sR)
    
End Function

'----------------------------------------------------------------
Public Function EvaluateExpression(sExpr As String) As String
'----------------------------------------------------------------
'TA 05/10/2001: Public function so that user need not reference goArezzo directly
' not yet used (I need to speak to Nicky)
'----------------------------------------------------------------
    
    EvaluateExpression = goArezzo.EvaluateExpression(sExpr)

End Function

'----------------------------------------------------------------
Public Function ResultOK(sResult As String) As String
'----------------------------------------------------------------
'TA 05/10/2001: Public function so that user need not reference goArezzo directly
' not yet used (I need to speak to Nicky)
'----------------------------------------------------------------
    
    ResultOK = goArezzo.ResultOK(sResult)

End Function

'----------------------------------------------------------------
Public Function ArezzoDateToDouble(sArezzoDate As String) As String
'----------------------------------------------------------------
'TA 05/10/2001: Public function so that user need not reference goArezzo directly
' not yet used (I need to speak to Nicky)
'----------------------------------------------------------------
    
    ArezzoDateToDouble = goArezzo.ArezzoDateToDouble(sArezzoDate)

End Function

'----------------------------------------------------------------
Public Function FormatDate(sArezzoDate As String, sFormat As String) As String
'----------------------------------------------------------------
'TA 05/10/2001: Public function so that user need not reference goArezzo directly
' not yet used (I need to speak to Nicky)
'----------------------------------------------------------------
    
    FormatDate = goArezzo.FormatDate(sArezzoDate, sFormat)

End Function

'----------------------------------------------------------------
Public Function ReadValidDate(ByVal sDateString As String, _
                                ByVal sFormatString As String, _
                                ByRef sArezzoDate As String) As String
'----------------------------------------------------------------
' NCJ - Read a valid date from sDateString according to the format given
' Returns correctly validated date string
' and sArezzoDate is the "actual" date value as an Arezzo term
' (see also ConvertDateFromArezzo)
' NCJ 25 Sep 01 - Updated for use in MACRO 2.2 (Leave this in!)
'----------------------------------------------------------------
    
    ReadValidDate = goArezzo.ReadValidDate(sDateString, sFormatString, sArezzoDate)

End Function

'----------------------------------------------------------------
Public Function ConvertDateFromArezzo(ByVal sArezzoDate As String) As Double
'----------------------------------------------------------------
' Convert from Arezzo date string
' to VB's internal format as double
' Returns 0 if sDate not a valid date
' Assume sArezzoDate is one of
'   "date(Y,M,D)"
'   "time(H,Mn,S)"
'   "date(Y,M,D,H,Mn,S)"
' NCJ 25 Sep 01 - Updated for use in MACRO 2.2 (Leave this in!)
'----------------------------------------------------------------

    ConvertDateFromArezzo = goArezzo.ArezzoDateToDouble(sArezzoDate)

End Function

'------------------------------------------------------------------------
Public Sub DealWithArezzoTasks(oArezzo As Arezzo_DM, colDone As Collection, _
                    ByRef bSomethingAlreadyDone As Boolean)
'------------------------------------------------------------------------
' NEW FOR MACRO 2.2
' Deal with any non-MACRO Arezzo tasks that need attention.
' This is a recursive routine,
' and colDone is a collection of TaskIds that we did last time so don't do them again
' (Should be Nothing on first top-level call).
' NCJ 12/10/01 - Allow an in_progress decision to come through again as permited
' TA 26/10/2001 - Arezzo now passed in (global one not used)
' NCJ 31 Jan 03 - Retrun result in bSomethingAlreadyDone
'---------------------------------------------------------------------
Dim colTasks As Collection
Dim bSomethingHappenedThisTime As Boolean
Dim oTask As TaskInstance
Dim colDoneThisTime As Collection
Dim colAlreadyDone As Collection

    ' Get the tasks to process
    Set colTasks = oArezzo.GetArezzoTasks
    
    If colTasks.Count = 0 Then Exit Sub
    
    ' Remember if we did anything this time
    bSomethingHappenedThisTime = False
    
    ' Initialise the ones we do in this call
    Set colDoneThisTime = New Collection
    ' Set up the ones already done before
    If colDone Is Nothing Then
        Set colAlreadyDone = New Collection
    Else
        Set colAlreadyDone = colDone
    End If
    
    For Each oTask In colTasks
        ' Check we haven't already done it
        If Not CollectionMember(colAlreadyDone, "K" & oTask.TaskKey, False) Then
            ' NCJ 31 Jan 03 - Added oArezzo argument to Display calls
            Select Case oTask.TaskType
            Case "decision"
                If oTask.TaskState = "permitted" Then
                    bSomethingHappenedThisTime = frmArezzoDecision.Display(oTask, oArezzo)
                Else
                    ' Assume in_progress
                    ' Use Enquiry form to request the data
                    bSomethingHappenedThisTime = frmArezzoEnquiry.Display(oTask, oArezzo)
                End If
                
            Case "action"
                bSomethingHappenedThisTime = frmArezzoAction.Display(oTask, oArezzo)
                
            Case "enquiry"
                ' NCJ 31 Jan 03 - Added oArezzo argument
                bSomethingHappenedThisTime = frmArezzoEnquiry.Display(oTask, oArezzo)
                
            Case Else
                ' We're not expecting other task types
            End Select
            ' remember that we've done this one this time round
            ' so we don't offer it again
'            If Not bSomethingHappenedThisTime Then
            colDoneThisTime.Add oTask.TaskKey, "K" & oTask.TaskKey
'            End If
        End If
    Next
    
    Set colAlreadyDone = Nothing
    Set colTasks = Nothing
    Set oTask = Nothing
    
    ' Accumulate previous results
    bSomethingAlreadyDone = bSomethingAlreadyDone Or bSomethingHappenedThisTime
    
    If bSomethingHappenedThisTime Then
        ' Raise events and go round again
        Call oArezzo.ALM.GuidelineInstance.RunEngine
        Call oArezzo.GenerateArezzoEvents
        Call DealWithArezzoTasks(oArezzo, colDoneThisTime, bSomethingAlreadyDone)
    End If
    
    Set colDoneThisTime = Nothing
    
End Sub
