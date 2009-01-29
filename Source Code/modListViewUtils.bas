Attribute VB_Name = "modListViewUtils"
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       modListViewUtils.bas
'   Author:     Paul Norris, 23/07/1999
'   Purpose:    Assorted listview utility functions.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   PN  30/09/99    Commented code more thoroughly
'   WillC   11/10/99 Added error handlers
'------------------------------------------------------------------------------------'
Option Explicit

' windows api declerations
Public Const LVM_FIRST As Long = &H1000
Public Const LVM_SETCOLUMNWIDTH As Long = LVM_FIRST + 30
Public Const LVS_EX_GRIDLINES As Long = &H1
Public Const LVS_EX_SUBITEMIMAGES As Long = &H2
Public Const LVS_EX_CHECKBOXES As Long = &H4
Public Const LVS_EX_TRACKSELECT As Long = &H8
Public Const LVS_EX_HEADERDRAGDROP As Long = &H10
Public Const LVS_EX_FULLROWSELECT As Long = &H20
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE As Long = LVM_FIRST + 55
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = LVM_FIRST + 54

' sorting enumerators
Public Enum LVItemTypes
   LVTDate = 0
   LVTNumber = 1
   LVTBinary = 2
   LVTAlphabetic = 3
End Enum

' windows api declerations
Public Enum ListSortOrderConstants
   LVTAscending = 0
   LVTDescending = 1
End Enum

' windows api declerations
'MLM 13/09/01: This is now defined in libListView, use that class instead.
'Public Enum LVSCW_Styles
'   LVSCW_AUTOSIZE = -1
'   LVSCW_AUTOSIZE_USEHEADER = -2
'End Enum

' windows api declerations
Public Enum LVStylesEx
   LVSTCheckBoxes = LVS_EX_CHECKBOXES
   LVSTFullRowSelect = LVS_EX_FULLROWSELECT
   LVSTGridlines = LVS_EX_GRIDLINES
   LVSTHeaderDragDrop = LVS_EX_HEADERDRAGDROP
   LVSTSubItemImages = LVS_EX_SUBITEMIMAGES
   LVSTTrackSelect = LVS_EX_TRACKSELECT
End Enum

' windows api declerations
Public Declare Function SendMessage Lib "user32" Alias _
    "SendMessageA" (ByVal hWnd As Long, _
    ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long

'----------------------------------------------------------------------------------------'
Public Sub SortListview(lv As Object, _
    iSortKey As Integer, _
    iSortOrder As ListSortOrderConstants, _
    iSortType As LVItemTypes)
'----------------------------------------------------------------------------------------'
' this function will dynamically add a column header to the list view with width = 0
' it will then add the date or number data to the column
' it will then sort the listview on the new column
' it will then remove the new column
'----------------------------------------------------------------------------------------'
    Dim oItem As ListItem       'For for..each loop
    Dim sFormat As String       'Format for sort column
    Dim iCols As Integer        'Number of columns
On Error GoTo ErrHandler
    
With lv
        
        'Get column count
        iCols = .ColumnHeaders.Count
        
        If iSortType <> LVTAlphabetic Then
            'We need to sort by date or number
            Select Case iSortType
            Case LVTDate
                'Set format to sort by date
                sFormat = "yyyymmddhhnnss"
            Case LVTNumber
                'Set format to sort by number
                sFormat = String(20, "0") & "." & String(20, "0")
            End Select
            
            'Lock updates to prevent flicker
            Call LockWindow(lv)
            
            'Add sort column
            .ColumnHeaders.Add
            
            'Increment column count
            iCols = iCols + 1
            
            If iSortKey = 0 Then
                'Sorting by first col so use item text
                For Each oItem In .ListItems
                    'Add formatted string to sort column
                    oItem.SubItems(iCols - 1) = Format(oItem.Text, sFormat)
                Next
            Else
                'Sorting by other col so use relevant subitem text
                For Each oItem In .ListItems
                    'Add formatted string to sort column
                    oItem.SubItems(iCols - 1) = Format(oItem.SubItems(iSortKey), sFormat)
                Next
            End If
            
            'Sort key is the hidden col
            iSortKey = iCols - 1
                
        End If
        
        'Set listview sort properties
        .SortOrder = iSortOrder
        .SortKey = iSortKey
        .Sorted = True
        
        If iSortType <> LVTAlphabetic Then
            'We need to remove the sort col and unlock
            'the window
            .ColumnHeaders.Remove iCols
            Call UnlockWindow
        End If
    End With
 
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "TrimNull", "modListViewUtils.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub
'----------------------------------------------------------------------------------------'
Public Function LVSetStyleEx(lv As Object, _
                            ByVal NewStyle As LVStylesEx, _
                            ByVal NewVal As Boolean) As Boolean
'----------------------------------------------------------------------------------------'
' set the listview style using the windows api
'----------------------------------------------------------------------------------------'
   Dim nStyle As Long
On Error GoTo ErrHandler

   ' get the current ListView style
   nStyle = SendMessage(lv.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, ByVal 0&)
   
   If NewVal Then
      ' set the extended style bit
      nStyle = nStyle Or NewStyle
   Else
      ' remove the extended style bit
      nStyle = nStyle Xor NewStyle
   End If
   
   ' set the new ListView style
   LVSetStyleEx = CBool(SendMessage(lv.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, ByVal nStyle))
 
Exit Function
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "LVSetStyleEx", "modListViewUtils.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Function
'----------------------------------------------------------------------------------------'
Public Sub LVSetColWidth(lv As Object, ByVal ColumnIndex As Long, ByVal Style As LVSCW_Styles)
'----------------------------------------------------------------------------------------'
' autoresize a column so no need to specify the column size in forms
'----------------------------------------------------------------------------------------'
On Error GoTo ErrHandler
   
   ' if you include the header in the sizing then the last column will
   ' automatically size to fill the remaining listview width.
   With lv
      ' verify that the listview is in report view and that the column exists
      If .View = lvwReport Then
         If ColumnIndex >= 1 And ColumnIndex <= .ColumnHeaders.Count Then
            Call SendMessage(.hWnd, LVM_SETCOLUMNWIDTH, ColumnIndex - 1, ByVal Style)
         End If
      End If
   End With
 
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "LVSetColWidth", "modListViewUtils.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub
'----------------------------------------------------------------------------------------'
Public Sub LVSetAllColWidths(lv As Object, ByVal Style As LVSCW_Styles)
'----------------------------------------------------------------------------------------'
' autoresize a column so no need to specify sizes in forms
'----------------------------------------------------------------------------------------'
On Error GoTo ErrHandler

Dim ColumnIndex As Long
On Error GoTo ErrHandler

   '--- loop through all of the columns in the listview and size each
   With lv
      For ColumnIndex = 1 To .ColumnHeaders.Count
         LVSetColWidth lv, ColumnIndex, Style
      Next ColumnIndex
   End With
 
Exit Sub
ErrHandler:
  Select Case MACROCodeErrorHandler(Err.Number, Err.Description, "LVSetAllColWidths", "modListViewUtils.bas")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
   End Select
    
End Sub
