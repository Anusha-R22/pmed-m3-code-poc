VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GridCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "GridCell object for MACRO 2.2 ScheduleGrid"
'----------------------------------------------------------------------------------------'
'   File:       GridCell.cls
'   Copyright:  InferMed Ltd. 2001-2004. All Rights Reserved
'   Author:     Ash Trebi-Ollennu, June 2001
'   Purpose:    GridCell object for MACRO 2.2 ScheduleGrid
'----------------------------------------------------------------------------------------'
' Revisions:
'   AT-O 29 June 01 - Initial Development
'   NCJ 3 July - Further work
'   NCJ 6 Jul 01 - Added VisitInst property
'   NCJ 24 Mar 04 - Kill VisitInst on Terminate
'----------------------------------------------------------------------------------------'
Option Explicit

Private moEFormInst As EFormInstance
Private moVisitInst As VisitInstance
Private mnCell As eCellType
Private moEForm As eFormRO
Private moVisit As VisitRO

' The type of a cell in the Schedule Grid
' (Blank means there's nothing in the cell,
'   Active means there's an eFormInstance in the cell,
'   Inactive means there's an inactive eForm in the cell)
' Note that eForm Header cells are also of type Blank
' but Visit header cells may be Active or Inactive
Public Enum eCellType
    Blank = 0
    Active = 1
    Inactive = 2
End Enum

Private mnRow As Integer
Private mnCol As Integer

'------------------------------------------------
Public Property Get Row() As Integer
Attribute Row.VB_Description = "This cells's row number."
'------------------------------------------------
'This cells's row number.
'------------------------------------------------
    
    Row = mnRow

End Property

'------------------------------------------------
Friend Property Let Row(nRow As Integer)
'------------------------------------------------
    
     mnRow = nRow

End Property

'------------------------------------------------
Public Property Get Col() As Integer
Attribute Col.VB_Description = "This cells's column number."
'------------------------------------------------
'This cells's column number.
'------------------------------------------------
    
    Col = mnCol

End Property

'------------------------------------------------
Friend Property Let Col(nCol As Integer)
'------------------------------------------------
    
     mnCol = nCol

End Property

'------------------------------------------------
Public Property Get CellType() As eCellType
Attribute CellType.VB_Description = "Whether Blank, Active eForm or Inactive eForm"
'------------------------------------------------
' Whether Blank, Active eForm or Inactive eForm
'------------------------------------------------
    
    CellType = mnCell

End Property

'------------------------------------------------
Public Property Let CellType(nNewCell As eCellType)
'------------------------------------------------
    
     mnCell = nNewCell

End Property

'---------------------------------------------------
Public Property Get eFormInst() As EFormInstance
Attribute eFormInst.VB_Description = "The cell's eForm instance (may be nothing)."
'---------------------------------------------------
' The cell's eForm instance (may be nothing).
'---------------------------------------------------
    
    Set eFormInst = moEFormInst

End Property

'---------------------------------------------------
Public Property Set eFormInst(oNewFormINST As EFormInstance)
'---------------------------------------------------
    
    Set moEFormInst = oNewFormINST

End Property

'---------------------------------------------------
Public Property Get VisitInst() As VisitInstance
Attribute VisitInst.VB_Description = "The cell's Visit instance (may be nothing)."
'---------------------------------------------------
' The cell's Visit instance (may be nothing).
'---------------------------------------------------
    
    Set VisitInst = moVisitInst

End Property

'---------------------------------------------------
Public Property Set VisitInst(oNewVisitInst As VisitInstance)
'---------------------------------------------------
    
    Set moVisitInst = oNewVisitInst

End Property

'--------------------------------------------
Public Property Get eForm() As eFormRO
Attribute eForm.VB_Description = "The cell's eForm"
'--------------------------------------------
' The cell's eForm
'--------------------------------------------
            
     Set eForm = moEForm

End Property

'--------------------------------------------
Public Property Set eForm(oNewEForm As eFormRO)
'--------------------------------------------
            
     Set moEForm = oNewEForm

End Property

'----------------------------------------------
Public Property Get Visit() As VisitRO
Attribute Visit.VB_Description = "The Visit definition object to which this cell belongs."
'----------------------------------------------
' The Visit definition object to which this cell belongs.
'----------------------------------------------

    Set Visit = moVisit

End Property

'----------------------------------------------
Public Property Set Visit(oNewVisit As VisitRO)
'----------------------------------------------

    Set moVisit = oNewVisit

End Property

'----------------------------------------------
Private Sub Class_Initialize()
'----------------------------------------------

    Set moVisit = Nothing
    Set moEForm = Nothing
    Set moEFormInst = Nothing
    Set moVisitInst = Nothing   ' Added NCJ 24 Mar 04

End Sub

'----------------------------------------------
Private Sub Class_Terminate()
'----------------------------------------------
' Tidy up before we go home
'----------------------------------------------

    Call Terminate

End Sub
 
'----------------------------------------------
Friend Sub Terminate()
'----------------------------------------------
' NCJ 25 Mar 04
'----------------------------------------------

    Set moVisit = Nothing
    Set moEForm = Nothing
    Set moEFormInst = Nothing
    Set moVisitInst = Nothing   ' Added NCJ 24 Mar 04

End Sub

 
 
