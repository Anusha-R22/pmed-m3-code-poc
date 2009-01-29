VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInputGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   File:       clsInputGrid.cls
'   Copyright:  InferMed Ltd. 2000-2006. All Rights Reserved
'   Author:     Toby Aldridge, November 2000
'   Purpose:    Class to handle input to a grid cell through a text box
'----------------------------------------------------------------------------------------'
' REVISIONS:
' TA 05/12/2000: Tabbing around grid implemented
'               (Make sure the tab index of the grid is one greater than of the text box for tabbing to work)
' TA 06/12/2000: problems with tabbing so code commented out
' NCJ 14 Jun 06 - Added CanEdit argument to Init
'----------------------------------------------------------------------------------------'

Option Explicit

Public WithEvents mMsFlexGrid As MSFlexGrid
Attribute mMsFlexGrid.VB_VarHelpID = -1
Public WithEvents mTextbox As TextBox
Attribute mTextbox.VB_VarHelpID = -1
Public mcmdButton As CommandButton
Attribute mcmdButton.VB_VarHelpID = -1

Private mlButtonCol As Long

'table to store data behind grid
Private mtblTable As clsDataTable

Public Event CellChange(lRow As Long, lCol As Long)
Public Event RowAdded()
Public Event RowDeleted(lRow As Long)

Private mbEditMode As Boolean

' NCJ 14 Jun 06 - Allow editing?
Private mbCanEdit As Boolean

'----------------------------------------------------------------------------------------'
Public Sub Init(tblTable As clsDataTable, Optional lButtonCol As Long = -1, _
            Optional bCanEdit As Boolean = True)
'----------------------------------------------------------------------------------------'
' NCJ 14 Jun 06 - Added bCanEdit
'----------------------------------------------------------------------------------------'
    
    Set mtblTable = tblTable.Duplicate
    mtblTable.InsertColumn 1, "      ", "        "
    
    mlButtonCol = lButtonCol
    mbCanEdit = bCanEdit
    
    mbEditMode = False
    mtblTable.Add RecordInit(mtblTable.Cols)
    TabletoGrid mMsFlexGrid, mtblTable
    mMsFlexGrid.FixedCols = 1
    mMsFlexGrid.Row = 1
    mMsFlexGrid.Col = 1
    mbEditMode = True
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub Class_Initialize()
'----------------------------------------------------------------------------------------'

    mbEditMode = True
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub mMsFlexGrid_EnterCell()
'----------------------------------------------------------------------------------------'

    ShowTextBox
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub mMsFlexGrid_GotFocus()
'----------------------------------------------------------------------------------------'

    ShowTextBox

End Sub

'----------------------------------------------------------------------------------------'
Private Sub ShowTextBox()
'----------------------------------------------------------------------------------------'
' NCJ 14 Jun 06 - Only if we can edit
'----------------------------------------------------------------------------------------'

    If mbEditMode And mbCanEdit Then
   
        With mMsFlexGrid
            mTextbox.Top = .Top + .CellTop
            mTextbox.Width = .CellWidth
            mTextbox.Left = .Left + .CellLeft
            mTextbox.Height = .CellHeight
            mTextbox.Text = .Text
            mTextbox.Visible = True
            If Not (mcmdButton Is Nothing) Then
                If .Col = mlButtonCol Then
                    mcmdButton.Top = mTextbox.Top - (mcmdButton.Height - mTextbox.Height) / 2
                    mcmdButton.Left = mTextbox.Left + mTextbox.Width
                    mcmdButton.Visible = True
                Else
                    mcmdButton.Visible = False
                End If
            End If
            mTextbox.SetFocus
        End With
        
    Else
    
        mTextbox.Visible = False
        If Not (mcmdButton Is Nothing) Then
            mcmdButton.Visible = False
        End If
        
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub mMsFlexGrid_Scroll()
'----------------------------------------------------------------------------------------'
'ensure textbox moves with cell
'----------------------------------------------------------------------------------------'

    ShowTextBox

End Sub

'----------------------------------------------------------------------------------------'
Private Sub mTextBox_Change()
'----------------------------------------------------------------------------------------'
Dim i As Long

    With mMsFlexGrid
        .Text = mTextbox.Text
        mtblTable(.Row, .Col + 1) = mTextbox.Text
        
        'notify of cell change
        RaiseEvent CellChange(.Row, .Col + 1)
        If .Row = mtblTable.Rows And Not RowEmpty(.Row) Then
            RowAdd
        End If
            
    End With

End Sub

'----------------------------------------------------------------------------------------'
Private Sub NextCell(KeyCode As Integer)
'----------------------------------------------------------------------------------------'
' move to next cell
'----------------------------------------------------------------------------------------'
Dim bEditMode As Boolean
Dim lRow As Long
Dim lCol As Long

    With mMsFlexGrid
    
        bEditMode = mbEditMode
        mbEditMode = False
        ShowTextBox
        mbEditMode = bEditMode
        
        lCol = .Col
        lRow = .Row
        
        Select Case KeyCode
        Case vbKeyLeft
            If .Col = 1 Then
                lCol = .Cols - 1
                lRow = lRow - 1
            Else
                lCol = lCol - 1
            End If
        Case vbKeyRight
            If .Col = .Cols - 1 Then
                lCol = 1
                lRow = lRow + 1
            Else
                lCol = .Col + 1
            End If
        Case vbKeyUp
            lRow = .Row - 1
        Case vbKeyDown
            lRow = lRow + 1
        End Select
        
        'if moved and not greater than number or row or in header
        If ((lRow >= 1) And (lRow <= .Rows - 1)) And Not ((.Row = lRow) And (.Col = lCol)) Then
            .Col = lCol
            .Row = lRow
        End If
             
    End With
    
End Sub

'----------------------------------------------------------------------------------------'
Private Sub mTextbox_KeyPress(KeyAscii As Integer)
'----------------------------------------------------------------------------------------'
'if enter pressed move to next cell
'----------------------------------------------------------------------------------------'

    If Not mbCanEdit Then Exit Sub
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        NextCell vbKeyRight
    End If
    
End Sub

'----------------------------------------------------------------------------------------'
Private Function RowEmpty(lRow As Long) As Boolean
'----------------------------------------------------------------------------------------'
Dim i As Long
Dim bNoDataInRow As Boolean

    RowEmpty = True
    For i = 2 To mtblTable.Cols
        RowEmpty = RowEmpty And mtblTable(lRow, i) = ""
    Next
    
End Function

'----------------------------------------------------------------------------------------'
Private Function RowAdd()
'----------------------------------------------------------------------------------------'

    mtblTable.Add RecordInit(mtblTable.Cols)
    mMsFlexGrid.AddItem String(mtblTable.Cols, vbTab)

    'notify row added
    RaiseEvent RowAdded

End Function

'----------------------------------------------------------------------------------------'
Private Function RowDelete(lRow As Long)
'----------------------------------------------------------------------------------------'

    mtblTable.Delete lRow
    mMsFlexGrid.RemoveItem lRow
    RaiseEvent RowDeleted(lRow)
    
End Function

'----------------------------------------------------------------------------------------'
Public Property Get Table() As clsDataTable
'----------------------------------------------------------------------------------------'
Dim tblTable As clsDataTable

    Set tblTable = mtblTable.Duplicate
    tblTable.Delete (tblTable.Rows)
    tblTable.DeleteColumn (1)
    
    Set Table = tblTable
    
End Property

'----------------------------------------------------------------------------------------'
Private Sub mTextbox_Validate(Cancel As Boolean)
'----------------------------------------------------------------------------------------'
'deal with tab being pressed (requires grid to be next tab stop after textbox)
' delete empty rows
'----------------------------------------------------------------------------------------'
Dim i As Long
    
    With mMsFlexGrid
        For i = 1 To mtblTable.Rows - 1
            If RowEmpty(i) Then
                RowDelete i
            End If
        Next
    End With
'    mbEditMode = False
'    NextCell vbKeyRight
'    mbEditMode = True
'
End Sub
