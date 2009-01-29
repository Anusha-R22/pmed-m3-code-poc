VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QGroupInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'----------------------------------------------------------------------------------------'
'   File:       QGroupInstance.cls
'   Copyright:  InferMed Ltd. 2001-2004. All Rights Reserved
'   Author:     Nicky Johns, November 2001
'   Purpose:    Question Group Instance object for MACRO 3.0
'----------------------------------------------------------------------------------------'
' Revisions:
'   NCJ 16-22 Nov 01 - Initial development
'   NCJ 2 Jan 02 - Allow setting of status to NotApplicable
'   NCJ 7 Jan 02 - Added Complete property
'   NCJ 14 Feb 02 - Treat requested as Missing in UpdateStatus
'   NCJ 18 Sept 02 - Store EFI TaskId rather than the object
'   NCJ 14 Nov 02 - No longer have a QGIState (QGroup instances not loaded/saved)
'   NCJ 13 Feb 03 - Ignore Unobtainable mandatory questions in QGI completeness
'   NCJ 1 Apr 03 - Bug fix to GetNumberOfRows
'   NCJ 29 Apr 03 - Check for saving of last row
'   NCJ 30 Apr 03 - Routine to create necessary new responses
'   NCJ 7 May 03 - Added LastRowWillBeSaved
'   NCJ 4 Mar 04 - Allow forcing of new QGroup row in CreateNewRow (necessary for Web DE)
'   NCJ 24/25 Mar 04 - Added Terminate event
'   NCJ 29 Mar 04 - Bug 2207 - If mandatory QGroup is Not Applicable, treat it as Complete
'----------------------------------------------------------------------------------------'

Private moSubject As StudySubject

' The group definition
Private moGroupDef As QGroupRO
' Store eForm Task ID rather than eForm itself
Private mlEFormTaskID As Long

' NCJ 14 Nov 02 - Status and no. of rows are now calculated rather than stored and loaded
Private mnStatus As Integer
Private mnRows As Integer

Option Explicit

'--------------------------------------------------------------------'
Friend Sub Init(oQGroup As QGroupRO, oEFI As EFormInstance)
'--------------------------------------------------------------------'
' NCJ 14 Nov 01 - Initialise this object
'--------------------------------------------------------------------'

    Set moGroupDef = oQGroup
    ' NCJ 18 Sept 02 - Store Subject and eFormTaskId instead of EFI
    mlEFormTaskID = oEFI.eFormTaskId
    Set moSubject = oEFI.VisitInstance.Subject
    
    mnRows = GetNumberOfRows
    
    Call UpdateStatus

End Sub

'--------------------------------------------------------------------'
Public Property Get QGroup() As QGroupRO
'--------------------------------------------------------------------'
' The Question Group definition
'--------------------------------------------------------------------'

    Set QGroup = moGroupDef

End Property

'--------------------------------------------------------------------'
Public Property Get Status() As Integer
'--------------------------------------------------------------------'
' The status of the group.
' See also SetNotApplicableStatus and UpdateStatus
'--------------------------------------------------------------------'

    Status = mnStatus

End Property

'--------------------------------------------------------------------'
Public Property Get Rows() As Integer
'--------------------------------------------------------------------'
' The number of rows in the group
'--------------------------------------------------------------------'

    Rows = mnRows

End Property

'--------------------------------------------------------------------'
Public Property Get QGroupID() As Long
'--------------------------------------------------------------------'
' The group id for the group
'--------------------------------------------------------------------'

    QGroupID = moGroupDef.QGroupID

End Property

'--------------------------------------------------------------------'
Friend Sub SetNotApplicableStatus(ByVal bNotApplicable As Boolean)
'--------------------------------------------------------------------'
' Switch status to/from NotApplicable depending on bNotApplicable
' NB This doesn't do anything to the responses within the group
'--------------------------------------------------------------------'

    If bNotApplicable Then
        ' Set to NotApplicable
        mnStatus = eStatus.NotApplicable
    Else
        ' Change from NotApplicable
        If mnStatus = eStatus.NotApplicable Then
            ' Temporarily change it to Missing
            ' (and assume there'll be an UpdateStatus some time later)
            mnStatus = eStatus.Missing
        End If
    End If

End Sub

'--------------------------------------------------------------------'
Friend Property Get Complete() As Boolean
'--------------------------------------------------------------------'
' Is the group instance "complete"?
' i.e. Have all mandatory questions in the min. no. of repeats been answered?
' NCJ 29 Mar 04 - If gp is Not Applicable, treat as Complete
'--------------------------------------------------------------------'
Dim nRow As Integer
Dim bComplete As Boolean

    On Error GoTo ErrLabel
    
    bComplete = True
    
    ' NCJ 29 Mar 04 - Ignore (treat as Complete) if status is Not Applicable
    If Me.Status <> eStatus.NotApplicable Then
    
        ' Must have min. no. of rows filled in
        If Me.Rows < moGroupDef.MinRepeats Then
            bComplete = False
    '    ElseIf Me.Status <> eStatus.NotApplicable Then
        Else
            ' Check for complete minimum no. of rows
            For nRow = 1 To moGroupDef.MinRepeats
                If Not IsComplete(nRow) Then
                    bComplete = False
                    Exit For
                End If
            Next
        End If
    
    End If
    
    Complete = bComplete
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|QGroupInstance.Complete"
    
End Property

'--------------------------------------------------------------------'
Friend Sub UpdateStatus()
'--------------------------------------------------------------------'
' Update the status of the group (but only if it's not NotApplicable)
' Rows beyond MinRepeats only contribute if they are not blank
' Code based on EFormInstance.UpdateStatus
' NCJ 14 Feb 02 - Need to treat "Requested" as "Missing" for purposes of group status calculation
' (e.g. for a new group on a form)
'--------------------------------------------------------------------'
Dim oEFI As EFormInstance
Dim oElement As eFormElementRO
Dim oResponse As Response
Dim nRow As Integer
Dim nMaxOptionalStatus As Integer
Dim nMaxNonOptionalStatus As Integer
Dim bInform As Boolean
Dim nStatus As Integer
Dim nRowCount As Integer

Dim nResponseStatus As Integer
Dim bAllRequested As Boolean

    On Error GoTo ErrLabel
    
    ' Leave it if it's NotApplicable
    If Me.Status = eStatus.NotApplicable Then
        Exit Sub
    End If
    
    Set oEFI = Me.EFI
    
    nStatus = eStatus.Success
    bInform = False     ' Initialise as no "Inform" status responses
    bAllRequested = True  ' We'll be looking for a non-requested one
    
    ' Initialise to "requested" (lowest value)
    nMaxNonOptionalStatus = eStatus.Requested
    nMaxOptionalStatus = eStatus.Requested
    
    ' Go to min. of current rows and minrepeats
    If Me.Rows < moGroupDef.MinRepeats Then
        nRowCount = Me.Rows
    Else
        nRowCount = moGroupDef.MinRepeats
    End If
    
    ' All rows less than MinRepeats count towards the status
    For nRow = 1 To nRowCount
        For Each oElement In moGroupDef.Elements
            Set oResponse = oEFI.Responses.ResponseByElement(oElement, nRow)
            If Not oResponse Is Nothing Then
                ' Exclude Inform status on this first pass
                If oResponse.Status <> eStatus.Inform Then
                    ' If requested use value of Missing
                    If oResponse.Status = eStatus.Requested Then
                        nResponseStatus = eStatus.Missing
                    Else
                        bAllRequested = False
                        nResponseStatus = oResponse.Status
                    End If
                    ' Deal with optional questions separately
                    If oElement.IsOptional Then
                        nMaxOptionalStatus = Max(nMaxOptionalStatus, nResponseStatus)
                    Else
                        nMaxNonOptionalStatus = Max(nMaxNonOptionalStatus, nResponseStatus)
                    End If
                Else
                    ' Response has Inform status
                    bInform = True
                    bAllRequested = False
                End If
            End If
        Next        ' oElement
    Next        ' nRow
    
    ' All rows above MinRepeats only count if they're not blank
    If Me.Rows > moGroupDef.MinRepeats Then
        For nRow = moGroupDef.MinRepeats + 1 To Me.Rows
            If Not IsBlank(nRow) Then
                For Each oElement In moGroupDef.Elements
                    Set oResponse = oEFI.Responses.ResponseByElement(oElement, nRow)
                    If Not oResponse Is Nothing Then
                        If oResponse.Status <> eStatus.Inform Then
                            ' If requested use value of Missing
                            If oResponse.Status = eStatus.Requested Then
                                nResponseStatus = eStatus.Missing
                            Else
                                bAllRequested = False
                                nResponseStatus = oResponse.Status
                            End If
                            If oElement.IsOptional Then
                                nMaxOptionalStatus = Max(nMaxOptionalStatus, nResponseStatus)
                            Else
                                nMaxNonOptionalStatus = Max(nMaxNonOptionalStatus, nResponseStatus)
                            End If
                        Else
                            ' Response has Inform status
                            bInform = True
                            bAllRequested = False
                        End If
                    End If
                Next      ' oElement
            End If   ' If row not blank
        Next    ' nRow
    End If
    
    ' Now sort out which value to use
    If bAllRequested Then
        ' All the group questions are still Requested
        nStatus = eStatus.Requested
    ElseIf nMaxOptionalStatus < nMaxNonOptionalStatus _
     Or (nMaxOptionalStatus = eStatus.Missing And nMaxNonOptionalStatus > eStatus.Requested) Then
        nStatus = nMaxNonOptionalStatus
    Else
        nStatus = nMaxOptionalStatus
    End If
     
    ' If status is Requested it means
    ' we didn't find any questions that weren't "Inform" on the first pass
    ' so check if there are any "Inform" questions
    If nStatus = eStatus.Requested And bInform Then
        nStatus = eStatus.Success
    End If
    
    mnStatus = nStatus
    
    Set oElement = Nothing
    Set oResponse = Nothing
    Set oEFI = Nothing

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|QGroupInstance.UpdateStatus"

End Sub

'--------------------------------------------------------------------'
Private Function IsComplete(nRow As Integer) As Boolean
'--------------------------------------------------------------------'
' Return whether a row is complete
' i.e. if all mandatory questions have been answered
'--------------------------------------------------------------------'
Dim oEFI As EFormInstance
Dim oElement As eFormElementRO
Dim oResponse As Response
Dim i As Integer
Dim bComplete As Boolean

    Set oEFI = Me.EFI
    bComplete = True
    i = 0
    Do While bComplete And i < moGroupDef.Elements.Count
        i = i + 1
        Set oElement = moGroupDef.Elements(i)
        If oElement.IsMandatory Then
            Set oResponse = oEFI.Responses.ResponseByElement(oElement, nRow)
            If oResponse Is Nothing Then
                ' Can't be complete if there's no response
                bComplete = False
            Else
                Select Case oResponse.Status
                ' NCJ 13 Feb 03 - Ignore mandatory Unobtainable
'                Case eStatus.Missing, eStatus.Requested, eStatus.Unobtainable
                Case eStatus.Missing, eStatus.Requested
                    ' If mandatory, row is not complete
                    bComplete = False
                Case Else
                End Select
            End If
        End If
    Loop

    IsComplete = bComplete
    
    Set oElement = Nothing
    Set oResponse = Nothing

End Function

'--------------------------------------------------------------------'
Private Function IsBlank(nRow As Integer) As Boolean
'--------------------------------------------------------------------'
' Return whether a row is all blank
' It is "blank" if it contains no non-derived values
'--------------------------------------------------------------------'
Dim oEFI As EFormInstance
Dim oElement As eFormElementRO
Dim oResponse As Response
Dim i As Integer
Dim bBlank As Boolean

    Set oEFI = Me.EFI
    bBlank = True
    i = 0
    Do While bBlank And i < moGroupDef.Elements.Count
        i = i + 1
        Set oElement = moGroupDef.Elements(i)
        ' Ignore derived items
        If oElement.DerivationExpr = "" Then
            Set oResponse = oEFI.Responses.ResponseByElement(oElement, nRow)
            If Not oResponse Is Nothing Then
                If oResponse.Value > "" Then
                    bBlank = False
                End If
            End If
        End If
    Loop
    
    IsBlank = bBlank
    
    Set oElement = Nothing
    Set oResponse = Nothing
    Set oEFI = Nothing
    
End Function

'--------------------------------------------------------------------'
Public Function CreateNewRow(Optional bForce As Boolean = False) As Boolean
'--------------------------------------------------------------------'
' Create a new row for this Group Instance if necessary
' Returns TRUE is a new row was created, or FALSE otherwise
' NCJ 4 Mar 04 - If bForce is TRUE, create a new row regardless (necessary for Web DE)
'--------------------------------------------------------------------'
Dim bComplete As Boolean
Dim i As Integer

    On Error GoTo ErrLabel
    
    CreateNewRow = False

    ' If we're up to the max, nothing to do (even if forcing)
    If Me.Rows = moGroupDef.MaxRepeats Then Exit Function
    
    ' NCJ 4 Mar 04 - Only check for completeness if not forcing
    If Not bForce Then
        ' If there are any blank or incomplete rows we do nothing
        ' NB A row can be Blank *and* Complete if it contains no mandatory questions
        ' so we check for Blank first
        For i = 1 To mnRows
            If IsBlank(i) Then Exit Function
            If Not IsComplete(i) Then Exit Function
        Next
    End If
    
    ' Otherwise if every row is non-blank and complete we add a new one
    ' Increment row count and create the Responses
    mnRows = mnRows + 1
    Call Me.EFI.GetRowResponses(moGroupDef, mnRows)
    CreateNewRow = True

Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|QGroupInstance.CreateNewRow"

End Function

'----------------------------------------------------
Public Property Get EFI() As EFormInstance
'----------------------------------------------------
' NCJ 18 Sept 02 - Return the QGroupinstance's EFI.
'----------------------------------------------------

    Set EFI = moSubject.eFIByTaskId(mlEFormTaskID)

End Property

'----------------------------------------------------
Private Function GetNumberOfRows() As Integer
'----------------------------------------------------
' Calculate the number of rows based on the max. cycle number
' of any response in the group
' If there are no responses (e.g. a new QGroup instance), set rows = group initial rows
' NCJ 1 Apr 03 - Can't rely on first question in the group because it might be new!
'----------------------------------------------------
Dim nMaxCycleNo As Integer
Dim bReachedMaxRow As Boolean
Dim bResponsesFound As Boolean

    On Error GoTo ErrLabel
    
    ' Default to initial no. of rows
    nMaxCycleNo = moGroupDef.InitialRows
    
    bReachedMaxRow = False
    bResponsesFound = False
    
    Do
        ' Check whether any responses exist for this row
        If ResponsesExist(nMaxCycleNo) Then
            nMaxCycleNo = nMaxCycleNo + 1
            bResponsesFound = True
        Else
            bReachedMaxRow = True
            ' Decrement back to the last row for which there were responses
            If bResponsesFound Then
                nMaxCycleNo = nMaxCycleNo - 1
            End If
        End If
    Loop Until bReachedMaxRow

    GetNumberOfRows = nMaxCycleNo
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|QGroupInstance.GetNumberOfRows"

End Function

'----------------------------------------------------
Private Function ResponsesExist(ByVal nRow As Integer) As Boolean
'----------------------------------------------------
' Returns TRUE if any response exists in the group for this row number
'----------------------------------------------------
Dim oElement As eFormElementRO
Dim bResponseFound As Boolean
Dim oEFI As EFormInstance
Dim oResponse As Response

    Set oEFI = Me.EFI
    bResponseFound = False
    
    ' Is there a response with this cycle no?
    For Each oElement In moGroupDef.Elements
        If oElement.QuestionId > 0 Then
            Set oResponse = oEFI.Responses.ResponseByElement(oElement, nRow)
            If Not oResponse Is Nothing Then
                bResponseFound = True
                ' Jump out because we've found an existing response with this cycle no.
                Exit For
            End If
        End If
    Next

    ResponsesExist = bResponseFound

    Set oResponse = Nothing
    Set oEFI = Nothing
    Set oElement = Nothing

End Function

'----------------------------------------------------
Friend Sub CheckSaveOfLastRow()
'----------------------------------------------------
' See if we should save the responses in last row of the group
' We don't want to save them if the last row is blank and beyond the min. repeats
' Sets the IgnoreMe property of each response as required
'----------------------------------------------------
Dim oResponse As Response
Dim oElement As eFormElementRO
Dim oEFI As EFormInstance

    On Error GoTo ErrLabel
    
    Set oEFI = Me.EFI
    
    If Not LastRowWillBeSaved Then
        ' We must prevent saving of responses in this row
        For Each oElement In moGroupDef.Elements
            If oElement.QuestionId > 0 Then
                Set oResponse = oEFI.Responses.ResponseByElement(oElement, mnRows)
                Call oResponse.IgnoreMe
            End If
        Next
    End If
    
    Set oResponse = Nothing
    Set oEFI = Nothing
    Set oElement = Nothing
    
Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|QGroupInstance.CheckSaveOfLastRow"

End Sub

'----------------------------------------------------
Friend Function LastRowWillBeSaved() As Boolean
'----------------------------------------------------
' NCJ 7 May 03
' Will the last row of the group be saved?
' Not if it's currently blank and it's beyond the min. repeats
'----------------------------------------------------
Dim oResponse As Response
Dim oElement As eFormElementRO
Dim oEFI As EFormInstance
Dim bNeedToSave As Boolean

    On Error GoTo ErrLabel
    
    Set oEFI = Me.EFI
    
    ' Check the last row is beyond min. repeats
    If mnRows <= moGroupDef.MinRepeats Then
        bNeedToSave = True
    Else
        bNeedToSave = False
        For Each oElement In moGroupDef.Elements
            ' Only consider non-derived questions
            If oElement.QuestionId > 0 And oElement.DerivationExpr = "" Then
                Set oResponse = oEFI.Responses.ResponseByElement(oElement, mnRows)
                If oResponse.Value > "" Or oResponse.SavedStatus <> eStatus.Requested Then
                    ' We've found a need-to-save response in the row
                    bNeedToSave = True
                    Exit For
                End If
            End If
        Next
    End If

    LastRowWillBeSaved = bNeedToSave
    
    Set oResponse = Nothing
    Set oEFI = Nothing
    Set oElement = Nothing
    
Exit Function
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|QGroupInstance.LastRowWillBeSaved"

End Function

'----------------------------------------------------
Friend Sub CreateNewResponses()
'----------------------------------------------------
' NCJ 30 Apr 03
' Create any new responses that are necessary
' e.g. new row may be required, or new questions may have been added
'----------------------------------------------------
Dim oEFI As EFormInstance
Dim i As Integer

    On Error GoTo ErrLabel
    
    Set oEFI = Me.EFI
    
    ' First of all create any new responses for existing rows
    For i = 1 To mnRows
        Call oEFI.GetRowResponses(moGroupDef, i)
    Next

    ' Now see if we need a new row (in case last blank row wasn't saved last time)
    Call CreateNewRow
    
    Set oEFI = Nothing

Exit Sub
ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|QGroupInstance.CreateNewResponses"

End Sub

'----------------------------------------------------
Private Sub Class_Terminate()
'----------------------------------------------------
' NCJ 24 Mar 04
'----------------------------------------------------

    Call Terminate

End Sub

'----------------------------------------------------
Friend Sub Terminate()
'----------------------------------------------------
' NCJ 25 Mar 04
'----------------------------------------------------

    Set moSubject = Nothing
    Set moGroupDef = Nothing

End Sub
