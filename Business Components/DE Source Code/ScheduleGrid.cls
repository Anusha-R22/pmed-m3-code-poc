VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScheduleGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ScheduleGrid object for Macro 2.2"
'----------------------------------------------------------------------------------------
'   File:       ScheduleGrid.cls
'   Copyright:  InferMed Ltd. 2001-2006. All Rights Reserved
'   Author:     Ash Trebi-Ollennu, June 2001
'   Purpose:    ScheduleGrid object for Macro 2.2, 3.0
'----------------------------------------------------------------------------------------
' Revisions:
'   AT-O 29 June 01 - Initial Develpoment
'   NCJ 3 Jul 01 - Further work
'   NCJ 6 Jul 01 - Rewrote Build to include Row 0 and Col 0 as eform/Visit header rows
'               and got rid of eFormRow objects
'   TA 21 Sep 01 - Schedule now copes with visits that have no instances
'   MLM 29/08/02: Changed Build() to cope with Visit eForms, which should not be shown in DE.
'   NCJ 24 Mar 04 - Explicitly clear object collections in Terminate
'   NCJ 27 Feb 06 - Don't do CollectionRemoveAll in Terminate because it's too slow!
'----------------------------------------------------------------------------------------
Option Explicit

' Collection of GridCell objects
Private mcolGridCells As Collection

' Collection of keys to retrieve cell by taskid
Private mcolTaskIdKey As Collection

' The number of eForm rows in the grid
Private mnRows As Integer
Private mnCols As Integer


'----------------------------------------------
Friend Sub Build(oSubject As StudySubject)
'----------------------------------------------
' Build the ScheduleGrid object for the given Study Subject
' Accessible only within MACROSchedule DLL (not outside the DLL)
'
' MLM 29/08/02: Cope with Visit eForms.
'----------------------------------------------
Dim nVCol As Integer   ' Visit column
Dim oEForm As eFormRO
Dim nFCycle As Integer
Dim oSVisit As ScheduleVisit
Dim oEFI As EFormInstance
Dim oGridCell As GridCell
Dim oVI As VisitInstance

    ' initialise collections of GridCells and eFormRows
    Set mcolGridCells = New Collection
    Set mcolTaskIdKey = New Collection
    
    ' First of all create row 0, containing the Visit column headers
    For nVCol = 1 To oSubject.ScheduleVisits.Count
        Set oSVisit = oSubject.ScheduleVisits(nVCol)
        If oSVisit.VisitInstance Is Nothing Then
            ' Visit is inactive
            Set oGridCell = CreateCell(0, nVCol, Inactive)
        Else
            ' Visit is active
            Set oGridCell = CreateCell(0, nVCol, Active)
            Set oGridCell.VisitInst = oSVisit.VisitInstance
        End If
        Set oGridCell.Visit = oSVisit.Visit
    Next nVCol
    
    'Now iterate through the study's eForms to create the rows
    mnRows = 0
    For Each oEForm In oSubject.eForms
        Select Case oSubject.IsAnInstance(oEForm)
        Case True
            ' There is at least one instance of this eForm
            ' Start at cycle 1
            nFCycle = 1
            ' Loop while there is an instance of eForm with this cycle number
            Do While oSubject.IsAnInstance(oEForm, nFCycle)
                ' New grid row
                mnRows = mnRows + 1
                ' Create the "header" cell (column = 0)
                Set oGridCell = CreateCell(mnRows, 0, Blank)
                Set oGridCell.eForm = oEForm
                ' Start in first visit column
                nVCol = 1
                ' Go through the visits
                For Each oSVisit In oSubject.ScheduleVisits
                    ' See if there is an instance of this visit
                    Set oVI = oSVisit.VisitInstance
                    If oVI Is Nothing Then
                        'no corresponding visitinstance so all forms are inactive
                        If oSVisit.Visit.IsInVisit(oEForm) Then
                            'the form is in the visit, but only draw an inactive form icon if it isn't a Visit eForm
                            Select Case oSVisit.Visit.VisitEFormByEForm(oEForm).EFormUse
                            Case eEFormUse.User
                                ' It's an inactive eForm
                                Set oGridCell = CreateCell(mnRows, nVCol, Inactive)
                                Set oGridCell.eForm = oEForm
                                Set oGridCell.Visit = oSVisit.Visit
                            Case Else
                                Set oGridCell = CreateCell(mnRows, nVCol, Blank)
                            End Select
                        Else
                            ' It's blank
                            Set oGridCell = CreateCell(mnRows, nVCol, Blank)
                        End If
                    Else 'there is a visit instance
                        If oSVisit.Visit.IsInVisit(oEForm) Then
                            'the form is in the visit
                            Select Case oSVisit.Visit.VisitEFormByEForm(oEForm).EFormUse
                            Case eEFormUse.User
                                ' See if there is an instance of this eForm
                                Set oEFI = oSVisit.VisitInstance.eFormInstanceByEFormCycle(oEForm, nFCycle)
                                If Not oEFI Is Nothing Then
                                    ' There's an eForm instance
                                    Set oGridCell = CreateCell(mnRows, nVCol, Active)
                                    Set oGridCell.eFormInst = oEFI
                                    
                                    ' add to task id collection for retieval by taskid
                                    mcolTaskIdKey.Add CellKey(mnRows, nVCol), "K" & Format(oEFI.eFormTaskId)
                                    
                                    Set oGridCell.eForm = oEForm
                                    Set oGridCell.Visit = oSVisit.Visit
                                    Set oGridCell.VisitInst = oSVisit.VisitInstance
                                Else
                                    ' We only create an inactive form for the FIRST cycle row
                                    If nFCycle = 1 Then
                                        ' It's an inactive eForm
                                        Set oGridCell = CreateCell(mnRows, nVCol, Inactive)
                                        Set oGridCell.eForm = oEForm
                                        Set oGridCell.Visit = oSVisit.Visit
                                    Else
                                        ' It's blank
                                        Set oGridCell = CreateCell(mnRows, nVCol, Blank)
                                    End If
                                End If
                            Case Else 'not a user eForm
                                Set oGridCell = CreateCell(mnRows, nVCol, Blank)
                            End Select
                        Else
                            ' the eForm is not in the Visit
                            Set oGridCell = CreateCell(mnRows, nVCol, Blank)
                        End If
                    
                    End If 'oVI Is Nothing
                    
                    ' Next visit column
                    nVCol = nVCol + 1
                    
                Next    ' ScheduleVisit
                ' Go on to next cycle no.
                nFCycle = nFCycle + 1
            Loop
            ' No more cycles
        
        Case False
            ' There are no user instances of this eForm
            If oEForm.HideIfInactive Or Not oEForm.AppearsInSchedule Then
                ' Ignore this row
            Else
                ' Create a row
                mnRows = mnRows + 1
                ' Create the "header" cell (column = 0)
                Set oGridCell = CreateCell(mnRows, 0, Blank)
                Set oGridCell.eForm = oEForm
                ' Start in first visit column
                nVCol = 1
                ' Go through the visits
                For Each oSVisit In oSubject.ScheduleVisits
                    If oSVisit.Visit.IsInVisit(oEForm) Then
                        'the form is in the visit, but only draw an inactive form icon if it isn't a Visit eForm
                        Select Case oSVisit.Visit.VisitEFormByEForm(oEForm).EFormUse
                        Case eEFormUse.User
                            ' It's an inactive eForm
                            Set oGridCell = CreateCell(mnRows, nVCol, Inactive)
                            Set oGridCell.eForm = oEForm
                            Set oGridCell.Visit = oSVisit.Visit
                        Case Else
                            Set oGridCell = CreateCell(mnRows, nVCol, Blank)
                        End Select
                    Else
                        ' It's blank
                        Set oGridCell = CreateCell(mnRows, nVCol, Blank)
                    End If
                    
                    ' Next visit column
                    nVCol = nVCol + 1
                Next    ' ScheduleVisit
            End If
        End Select
        
    Next  ' eForm

    'sote the number of cols
    mnCols = oSubject.ScheduleVisits.Count

End Sub

'----------------------------------------------------------------------
Public Property Get Cells(ByVal nRow As Integer, ByVal nCol As Integer) As GridCell
Attribute Cells.VB_Description = "Pick up the GridCell in this row and column."
'----------------------------------------------------------------------
' Pick up the GridCell in this row and column.
'----------------------------------------------------------------------
    
    Set Cells = mcolGridCells.Item(CellKey(nRow, nCol))

End Property

'----------------------------------------------------------------------
Public Function CellByEFITaskId(ByVal lTaskId As Long) As GridCell
Attribute CellByEFITaskId.VB_Description = "Pick up the GridCell for this eFormTaskId."
'----------------------------------------------------------------------
'Pick up the GridCell for this eFormTaskId.
'----------------------------------------------------------------------
    
    Set CellByEFITaskId = mcolGridCells.Item(mcolTaskIdKey.Item("K" & Format(lTaskId)))

End Function

'-----------------------------------------------------
Public Property Get RowMax() As Integer
Attribute RowMax.VB_Description = "The max row number in the grid. This is calculated it during the Build process. The count is RowMax + 1 if you include Row 0 (the visit header row)."
'-----------------------------------------------------
' The max row number in the grid.
' This is calculated it during the Build process.
' The count is RowMax + 1 if you include Row 0 (the visit header row).
'-----------------------------------------------------

    RowMax = mnRows

End Property

'------------------------------------------------------
Public Property Get ColMax() As Integer
Attribute ColMax.VB_Description = "The max column number in the grid (same as number of ScheduleVisits). The count is ColMax + 1 if you include Column 0 (the eForm header column)."
'------------------------------------------------------
' The max column number in the grid
' (same as number of ScheduleVisits).
' The count is ColMax + 1 if you include Column 0 (the eForm header column).
'-----------------------------------------------------
    
        ColMax = mnCols

End Property

'--------------------------------------------------------------
Private Function CellKey(ByVal nRow As Integer, ByVal nCol As Integer) As String
'--------------------------------------------------------------
' The key for the cell at the given row and column
'--------------------------------------------------------------
    
    CellKey = nRow & "|" & nCol

End Function

'--------------------------------------------------------------
Private Function CreateCell(ByVal nRow As Integer, ByVal nCol As Integer, _
                    ByVal nCelltype As eCellType) As GridCell
'--------------------------------------------------------------
' Create new grid cell and add to collection
'-----------------------------------------------------------
Dim oCell As GridCell

    ' Create new GridCell
    Set oCell = New GridCell
    ' Set its type
    oCell.CellType = nCelltype
    'set col and row
    oCell.Row = nRow
    oCell.Col = nCol
    ' Add it to the collection
    mcolGridCells.Add oCell, CellKey(nRow, nCol)
    ' Return the object we just created
    Set CreateCell = oCell
    
    ' Tidy up before we go
    Set oCell = Nothing
    
End Function

'-----------------------------------------------------------
Private Sub Class_Initialize()
'-----------------------------------------------------------
' Set ourselves up
'-----------------------------------------------------------

    Set mcolGridCells = New Collection
    Set mcolTaskIdKey = New Collection
    mnRows = 0
    
End Sub

'-----------------------------------------------------------
Private Sub Class_Terminate()
'-----------------------------------------------------------
' Tidy up before we go home
'-----------------------------------------------------------

    Call Terminate
    
End Sub

'----------------------------------------------
Friend Sub Terminate()
'----------------------------------------------
' NCJ 25 Mar 04
' NCJ 27 Feb 06 - Don't do CollectionRemoveAll because it takes AGES for a large Schedule!
'(approx. 1 sec per 1000 cells - Roche's study has over 29000 cells)
'----------------------------------------------
Dim oGridCell As GridCell

    If Not mcolGridCells Is Nothing Then
        For Each oGridCell In mcolGridCells
            Call oGridCell.Terminate
        Next
' NCJ 27 Feb 06 - Don't do CollectionRemoveAll because it takes too long
'        Call CollectionRemoveAll(mcolGridCells)     ' NCJ 24 Mar 04
    End If
    Set mcolGridCells = Nothing
    Set oGridCell = Nothing
    
    Set mcolTaskIdKey = Nothing     ' Not an object collection

End Sub

 
 
 
 
